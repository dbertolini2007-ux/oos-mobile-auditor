import os, io, re, json, tempfile, base64, mimetypes
from datetime import date, timedelta
from difflib import SequenceMatcher
import streamlit as st
import pandas as pd
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="OOS Auditor", page_icon="📦", layout="centered")

st.markdown("""
<style>
.block-container {max-width: 820px; padding-top: 1rem; padding-bottom: 3rem;}
div[data-testid="stFileUploader"] {border: 2px dashed #777; border-radius: 14px; padding: 8px;}
.stButton button, .stDownloadButton button {min-height: 52px; font-size: 18px; border-radius: 12px; width: 100%;}
[data-testid="stMetricValue"] {font-size: 28px;}
@media (max-width: 640px) {.block-container {padding-left:.8rem;padding-right:.8rem;} h1{font-size:1.8rem!important;}}
</style>
""", unsafe_allow_html=True)

DARK = colors.HexColor("#2B2B2B")
LIGHT = colors.HexColor("#F2F2F2")
MID = colors.HexColor("#D9D9D9")


def norm(x):
    x = re.sub(r"[^a-z0-9]+", " ", str(x or "").lower())
    return " ".join(x.split())


def similarity(a, b):
    a, b = norm(a), norm(b)
    if not a or not b: return 0.0
    if a in b or b in a: return .95
    return SequenceMatcher(None, a, b).ratio()


def parse_json(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


def upload_for_model(client, uploaded):
    suffix = os.path.splitext(uploaded.name)[1] or ".bin"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
        f.write(uploaded.getvalue()); p = f.name
    try:
        with open(p, "rb") as fh:
            return client.files.create(file=fh, purpose="user_data").id
    finally:
        try: os.unlink(p)
        except: pass


def verified_outs(uploaded, client, model):
    # Excel/CSV never need the AI API. Read Column A directly.
    n = uploaded.name.lower()
    raw = uploaded.getvalue()
    if n.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(raw), header=None)
        return [str(v).strip() for v in df.iloc[:, 0].dropna() if str(v).strip()]
    if n.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw), header=None)
        return [str(v).strip() for v in df.iloc[:, 0].dropna() if str(v).strip()]

    instruction = (
        'Read this Verified Outs sheet. Anything in the LEFT COLUMN / Column A is a verified OOS item. '
        'Ignore Column B and every other column, including notes/status. Return JSON only in exactly this shape: '
        '{"verified_outs":["item 1","item 2"]}'
    )

    # Images must be sent as image input, not as a generic input_file.
    if n.endswith((".png", ".jpg", ".jpeg")):
        mime = "image/png" if n.endswith(".png") else "image/jpeg"
        data_url = f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"
        r = client.responses.create(
            model=model,
            input=[{"role":"user","content":[
                {"type":"input_image","image_url":data_url},
                {"type":"input_text","text":instruction}
            ]}]
        )
    elif n.endswith(".pdf"):
        # PDFs use the Files API + input_file.
        fid = upload_for_model(client, uploaded)
        r = client.responses.create(
            model=model,
            input=[{"role":"user","content":[
                {"type":"input_file","file_id":fid},
                {"type":"input_text","text":instruction}
            ]}]
        )
    else:
        raise ValueError("Unsupported Verified Outs file type. Use XLSX, XLS, CSV, PDF, PNG, JPG, or JPEG.")

    parsed = parse_json(r.output_text)
    return [str(v).strip() for v in parsed.get("verified_outs", []) if str(v).strip()]


def pick_data(uploaded, client, model):
    fid = upload_for_model(client, uploaded)
    prompt = """Read these vending/micro-market pick sheets carefully.
Return JSON only:
{"orders":[{"account":"name","picker":"name if visible","pages":[1],"total_ordered":123,
"os_lines":[{"product":"full product name","ordered_qty":5,"page":1,"confidence":0.98}]}]}
Rules:
- Consolidate continuation pages for the same account/order.
- Determine total ordered units without double-counting cumulative totals printed on continuation pages.
- Count an item as OOS ONLY if the fills/handwritten area explicitly says OS or unmistakably indicates OS.
- Do not treat blanks, checkmarks, partial-fill numbers, or unrelated notes as OS.
- ordered_qty is the quantity ordered for the OS line, not the partial fill.
- confidence below 0.85 when handwriting/product identification is uncertain.
- Preserve product distinctions: e.g. Cheetos Flamin Hot and Chester Flamin Hot Fries are different products.
"""
    r = client.responses.create(model=model, input=[{"role":"user","content":[{"type":"input_file","file_id":fid},{"type":"input_text","text":prompt}]}])
    return parse_json(r.output_text)


def classify(data, verified, threshold):
    rows = []
    for order in data.get("orders", []):
        for x in order.get("os_lines", []):
            prod = x.get("product", "")
            candidates = [(v, similarity(prod, v)) for v in verified]
            _, score = max(candidates, key=lambda z:z[1]) if candidates else ("", 0)
            conf = float(x.get("confidence", 0) or 0)
            if conf < .85: status = "NEEDS REVIEW"
            elif score >= threshold: status = "VERIFIED OOS"
            else: status = "FALSE OOS / PICKING EXCEPTION"
            rows.append({"Account":order.get("account","Unknown"), "Picker":order.get("picker","") or "",
                         "Product":prod, "Qty":int(x.get("ordered_qty",0) or 0), "Page":x.get("page",""),
                         "Confidence":round(conf,2), "Status":status})
    return pd.DataFrame(rows)


def table_style(font=8):
    return TableStyle([
        ("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),font),
        ("GRID",(0,0),(-1,-1),.35,colors.HexColor("#AAAAAA")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[colors.white,LIGHT]),
        ("BACKGROUND",(0,-1),(-1,-1),MID),("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold")])


def add_table(story, headers, rows, widths, font=8, total_row=True):
    t = Table([headers] + rows, colWidths=widths, repeatRows=1)
    ts = table_style(font)
    if not total_row:
        ts.add("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, LIGHT])
        ts.add("BACKGROUND", (0,-1), (-1,-1), colors.white)
        ts.add("FONTNAME", (0,-1), (-1,-1), "Helvetica")
    t.setStyle(ts); story.append(t)


def summary_kpis(story, total_ordered, os_units, false_units, false_lines):
    pct = os_units / total_ordered * 100 if total_ordered else 0
    k = [["TOTAL UNITS ORDERED","OOS UNITS","OOS %","FALSE OOS UNITS"],
         [f"{total_ordered:,}",f"{os_units:,}",f"{pct:.2f}%",f"{false_units:,}"]]
    t = Table(k, colWidths=[1.75*inch]*4)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                           ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),
                           ("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),("FONTSIZE",(0,1),(-1,1),15),
                           ("BOX",(0,0),(-1,-1),.5,colors.grey),("INNERGRID",(0,0),(-1,-1),.3,colors.grey),
                           ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)]))
    story += [t, Spacer(1,14)]


def analysis_pdf(data, df, title="Out-of-Stock Analysis", subtitle="Only pick-sheet lines explicitly marked OS are counted."):
    b = io.BytesIO(); styles = getSampleStyleSheet()
    h2 = ParagraphStyle("h2x", parent=styles["Heading2"], fontSize=15, leading=18, spaceBefore=8, spaceAfter=7)
    doc = SimpleDocTemplate(b, pagesize=landscape(letter), leftMargin=22, rightMargin=22, topMargin=22, bottomMargin=22)
    story = [Paragraph(title, styles["Title"]), Paragraph(subtitle, ParagraphStyle("sub",parent=styles["BodyText"],alignment=TA_CENTER)), Spacer(1,12)]
    total_ordered = sum(int(o.get("total_ordered",0) or 0) for o in data.get("orders",[]))
    os_units = int(df.Qty.sum()) if len(df) else 0
    false_df = df[df.Status=="FALSE OOS / PICKING EXCEPTION"] if len(df) else df
    false_units = int(false_df.Qty.sum()) if len(false_df) else 0
    summary_kpis(story,total_ordered,os_units,false_units,len(false_df))

    story.append(Paragraph("1. Out-of-Stock Percentage by Location", h2))
    loc=[]
    for o in data.get("orders",[]):
        ordered=int(o.get("total_ordered",0) or 0); oq=sum(int(x.get("ordered_qty",0) or 0) for x in o.get("os_lines",[]))
        loc.append([o.get("account","Unknown"),ordered,oq,f"{(oq/ordered*100 if ordered else 0):.2f}%"])
    loc.sort(key=lambda r: float(r[3][:-1]), reverse=True)
    loc.append(["TOTAL",total_ordered,os_units,f"{(os_units/total_ordered*100 if total_ordered else 0):.2f}%"])
    add_table(story,["Location","Total Units Ordered","Units on OOS Lines","OOS %"],loc,[3.5*inch,1.55*inch,1.55*inch,1.0*inch])
    story += [Spacer(1,16), Paragraph("2. OOS by Individual Item Quantity", h2)]

    if len(df):
        q = df.groupby("Product",dropna=False).agg(OS_Units=("Qty","sum"), Locations=("Account","nunique")).reset_index().sort_values(["OS_Units","Product"],ascending=[False,True])
        qrows=[]
        for i,r in enumerate(q.itertuples(index=False),1):
            qrows.append([i,r.Product,int(r.OS_Units),int(r.Locations),f"{(r.OS_Units/os_units*100 if os_units else 0):.1f}%"])
        qrows.append(["","TOTAL",os_units,"", "100.0%" if os_units else "0.0%"])
    else: qrows=[["","No OOS items",0,0,"0.0%"]]
    add_table(story,["Rank","Product","OOS Units","Locations OOS","% of OOS Units"],qrows,[.55*inch,4.4*inch,1.0*inch,1.0*inch,1.0*inch],7.5)

    story += [Spacer(1,16), Paragraph("3. OOS by Item Frequency / Occurrences", h2)]
    if len(df):
        f = df.groupby("Product",dropna=False).agg(Occurrences=("Product","size"),Locations=("Account","nunique"),OS_Units=("Qty","sum")).reset_index().sort_values(["Occurrences","OS_Units"],ascending=[False,False])
        frows=[[i,r.Product,int(r.Occurrences),int(r.Locations),int(r.OS_Units)] for i,r in enumerate(f.itertuples(index=False),1)]
        frows.append(["","TOTAL",len(df),"",os_units])
    else: frows=[["","No OOS items",0,0,0]]
    add_table(story,["Rank","Product","OOS Occurrences","Locations OOS","OOS Units"],frows,[.55*inch,4.4*inch,1.1*inch,1.0*inch,1.0*inch],7.5)
    doc.build(story); return b.getvalue()


def repick_pdf(df):
    b=io.BytesIO(); styles=getSampleStyleSheet()
    doc=SimpleDocTemplate(b,pagesize=letter,leftMargin=30,rightMargin=30,topMargin=28,bottomMargin=28)
    h=ParagraphStyle("loc",parent=styles["Heading2"],fontSize=12,leading=14,spaceAfter=5)
    story=[Paragraph("FALSE OOS — REPICK LIST",styles["Title"]),
           Paragraph("Items marked OOS on the pick sheet but not present in Column A of Verified Outs.",styles["BodyText"]),Spacer(1,12)]
    f=df[df.Status=="FALSE OOS / PICKING EXCEPTION"] if len(df) else df
    if not len(f): story.append(Paragraph("No false OOS items found.",styles["Heading2"]))
    else:
        for account,g in f.groupby("Account",sort=False):
            pickers=[str(x) for x in g.Picker.dropna().unique() if str(x).strip() and str(x).lower()!="nan"]
            picker=", ".join(pickers)
            header=f"{account}" + (f" &nbsp;&nbsp;&nbsp; Picker: {picker}" if picker else "")
            rows=[[int(r.Qty),r.Product] for r in g.itertuples(index=False)]
            t=Table([["QTY","PRODUCT"]]+rows,colWidths=[.75*inch,6.3*inch],repeatRows=1)
            t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                                   ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),.45,colors.grey),
                                   ("FONTSIZE",(0,0),(-1,-1),9),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
                                   ("ALIGN",(0,1),(0,-1),"CENTER"),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,LIGHT])]))
            story.append(KeepTogether([Paragraph(header,h),t,Spacer(1,22)]))  # ~two blank lines between locations
    doc.build(story); return b.getvalue()


def current_history(report_date, data, df):
    ordered_map={str(o.get("account","Unknown")):int(o.get("total_ordered",0) or 0) for o in data.get("orders",[])}
    rows=[]
    if len(df):
        for r in df.itertuples(index=False):
            rows.append({"Date":str(report_date),"Account":r.Account,"Picker":r.Picker,"Product":r.Product,"Qty":int(r.Qty),"Status":r.Status,"OrderedUnits":ordered_map.get(str(r.Account),0)})
    # Preserve locations with zero OOS so WTD ordered-unit totals stay correct.
    present=set(df.Account.astype(str)) if len(df) else set()
    for account,ordered in ordered_map.items():
        if account not in present:
            rows.append({"Date":str(report_date),"Account":account,"Picker":"","Product":"","Qty":0,"Status":"NO OOS","OrderedUnits":ordered})
    return pd.DataFrame(rows)


def combine_history(prior_file, today_df):
    parts=[]
    if prior_file is not None:
        try: parts.append(pd.read_csv(io.BytesIO(prior_file.getvalue())))
        except Exception as e: st.warning(f"Could not read prior history CSV: {e}")
    parts.append(today_df)
    h=pd.concat(parts,ignore_index=True) if parts else today_df.copy()
    h["Date"]=pd.to_datetime(h["Date"],errors="coerce").dt.date
    for c in ["Qty","OrderedUnits"]: h[c]=pd.to_numeric(h[c],errors="coerce").fillna(0).astype(int)
    # Remove duplicate rows if the same daily export is uploaded again.
    return h.drop_duplicates().reset_index(drop=True)


def wtd_pdf(history, report_date):
    monday=report_date-timedelta(days=report_date.weekday())
    h=history[(history.Date>=monday)&(history.Date<=report_date)].copy()
    h=h[pd.to_datetime(h.Date).dt.weekday<5]
    b=io.BytesIO(); styles=getSampleStyleSheet(); h2=ParagraphStyle("h2w",parent=styles["Heading2"],fontSize=15,spaceBefore=8,spaceAfter=7)
    doc=SimpleDocTemplate(b,pagesize=landscape(letter),leftMargin=22,rightMargin=22,topMargin=22,bottomMargin=22)
    story=[Paragraph("Week-to-Date Out-of-Stock Analysis",styles["Title"]),Paragraph(f"Monday {monday:%m/%d/%Y} through {report_date:%m/%d/%Y} (Monday-Friday workweek)",ParagraphStyle("subw",parent=styles["BodyText"],alignment=TA_CENTER)),Spacer(1,12)]
    # OrderedUnits repeats for every OOS row; count once per Date+Account.
    ord_units=int(h[["Date","Account","OrderedUnits"]].drop_duplicates(["Date","Account"])["OrderedUnits"].sum()) if len(h) else 0
    os_h=h[(h.Product.fillna("").astype(str).str.strip()!="") & (h.Status!="NO OOS")]
    os_units=int(os_h.Qty.sum()) if len(os_h) else 0
    false_h=os_h[os_h.Status=="FALSE OOS / PICKING EXCEPTION"]
    summary_kpis(story,ord_units,os_units,int(false_h.Qty.sum()) if len(false_h) else 0,len(false_h))
    story.append(Paragraph("1. Week-to-Date OOS by Location",h2))
    locrows=[]
    for account,g in h.groupby("Account"):
        ordered=int(g[["Date","Account","OrderedUnits"]].drop_duplicates(["Date","Account"])["OrderedUnits"].sum())
        og=g[(g.Product.fillna("").astype(str).str.strip()!="") & (g.Status!="NO OOS")]
        oq=int(og.Qty.sum())
        locrows.append([account,ordered,oq,f"{(oq/ordered*100 if ordered else 0):.2f}%"])
    locrows.sort(key=lambda r:float(r[3][:-1]),reverse=True); locrows.append(["TOTAL",ord_units,os_units,f"{(os_units/ord_units*100 if ord_units else 0):.2f}%"])
    add_table(story,["Location","WTD Units Ordered","WTD OOS Units","WTD OOS %"],locrows,[3.5*inch,1.55*inch,1.55*inch,1.0*inch])
    story += [Spacer(1,16),Paragraph("2. Week-to-Date OOS by Individual Item Quantity",h2)]
    if len(os_h):
        q=os_h.groupby("Product").agg(OS_Units=("Qty","sum"),Locations=("Account","nunique")).reset_index().sort_values("OS_Units",ascending=False)
        qr=[[i,r.Product,int(r.OS_Units),int(r.Locations),f"{(r.OS_Units/os_units*100 if os_units else 0):.1f}%"] for i,r in enumerate(q.itertuples(index=False),1)]; qr.append(["","TOTAL",os_units,"","100.0%"])
    else: qr=[["","No OOS items",0,0,"0.0%"]]
    add_table(story,["Rank","Product","WTD OOS Units","Locations OOS","% of WTD OOS"],qr,[.55*inch,4.4*inch,1.0*inch,1.0*inch,1.0*inch],7.5)
    story += [Spacer(1,16),Paragraph("3. Week-to-Date OOS by Item Frequency / Occurrences",h2)]
    if len(os_h):
        f=os_h.groupby("Product").agg(Occurrences=("Product","size"),Locations=("Account","nunique"),OS_Units=("Qty","sum")).reset_index().sort_values(["Occurrences","OS_Units"],ascending=False)
        fr=[[i,r.Product,int(r.Occurrences),int(r.Locations),int(r.OS_Units)] for i,r in enumerate(f.itertuples(index=False),1)]; fr.append(["","TOTAL",len(os_h),"",os_units])
    else: fr=[["","No OOS items",0,0,0]]
    add_table(story,["Rank","Product","WTD OOS Occurrences","Locations OOS","WTD OOS Units"],fr,[.55*inch,4.3*inch,1.2*inch,1.0*inch,1.0*inch],7.5)
    doc.build(story); return b.getvalue()


def excel_export(report_date, data, df, history):
    """Create one polished Excel workbook containing the daily audit and Monday-Friday WTD analysis."""
    monday = report_date - timedelta(days=report_date.weekday())
    h = history[(history.Date >= monday) & (history.Date <= report_date)].copy()
    h = h[pd.to_datetime(h.Date).dt.weekday < 5]

    total_ordered = sum(int(o.get("total_ordered", 0) or 0) for o in data.get("orders", []))
    os_units = int(df.Qty.sum()) if len(df) else 0
    false_df = df[df.Status == "FALSE OOS / PICKING EXCEPTION"].copy() if len(df) else df.copy()
    false_units = int(false_df.Qty.sum()) if len(false_df) else 0
    os_pct = os_units / total_ordered if total_ordered else 0

    # Daily location summary
    loc_rows = []
    for o in data.get("orders", []):
        account = o.get("account", "Unknown")
        ordered = int(o.get("total_ordered", 0) or 0)
        oq = sum(int(x.get("ordered_qty", 0) or 0) for x in o.get("os_lines", []))
        loc_rows.append({"Location": account, "Total Units Ordered": ordered, "OOS Units": oq, "OOS %": oq / ordered if ordered else 0})
    daily_loc = pd.DataFrame(loc_rows)
    if len(daily_loc): daily_loc = daily_loc.sort_values(["OOS %", "OOS Units"], ascending=[False, False]).reset_index(drop=True)

    # Daily item quantity and frequency summaries
    if len(df):
        item_qty = df.groupby("Product", dropna=False).agg(**{"OOS Units": ("Qty", "sum"), "Locations Affected": ("Account", "nunique")}).reset_index()
        item_qty["% of All OOS Units"] = item_qty["OOS Units"] / os_units if os_units else 0
        item_qty = item_qty.sort_values(["OOS Units", "Product"], ascending=[False, True]).reset_index(drop=True)
        item_qty.insert(0, "Rank", range(1, len(item_qty) + 1))
        item_freq = df.groupby("Product", dropna=False).agg(**{"OOS Occurrences": ("Product", "size"), "Locations Affected": ("Account", "nunique"), "OOS Units": ("Qty", "sum")}).reset_index()
        item_freq = item_freq.sort_values(["OOS Occurrences", "OOS Units", "Product"], ascending=[False, False, True]).reset_index(drop=True)
        item_freq.insert(0, "Rank", range(1, len(item_freq) + 1))
    else:
        item_qty = pd.DataFrame(columns=["Rank", "Product", "OOS Units", "Locations Affected", "% of All OOS Units"])
        item_freq = pd.DataFrame(columns=["Rank", "Product", "OOS Occurrences", "Locations Affected", "OOS Units"])

    repick = false_df[[c for c in ["Account", "Picker", "Product", "Qty"] if c in false_df.columns]].copy() if len(false_df) else pd.DataFrame(columns=["Account", "Picker", "Product", "Qty"])
    repick = repick.rename(columns={"Account": "Location", "Qty": "Repick Qty"})

    # WTD calculations
    wtd_ordered = int(h[["Date", "Account", "OrderedUnits"]].drop_duplicates(["Date", "Account"])["OrderedUnits"].sum()) if len(h) else 0
    wtd_os = h[(h.Product.fillna("").astype(str).str.strip() != "") & (h.Status != "NO OOS")].copy() if len(h) else h.copy()
    wtd_os_units = int(wtd_os.Qty.sum()) if len(wtd_os) else 0
    wtd_false = wtd_os[wtd_os.Status == "FALSE OOS / PICKING EXCEPTION"].copy() if len(wtd_os) else wtd_os.copy()

    wtd_loc_rows = []
    if len(h):
        for account, g in h.groupby("Account"):
            ordered = int(g[["Date", "Account", "OrderedUnits"]].drop_duplicates(["Date", "Account"])["OrderedUnits"].sum())
            og = g[(g.Product.fillna("").astype(str).str.strip() != "") & (g.Status != "NO OOS")]
            oq = int(og.Qty.sum())
            wtd_loc_rows.append({"Location": account, "WTD Units Ordered": ordered, "WTD OOS Units": oq, "WTD OOS %": oq / ordered if ordered else 0})
    wtd_loc = pd.DataFrame(wtd_loc_rows)
    if len(wtd_loc): wtd_loc = wtd_loc.sort_values(["WTD OOS %", "WTD OOS Units"], ascending=[False, False]).reset_index(drop=True)

    if len(wtd_os):
        wtd_item = wtd_os.groupby("Product", dropna=False).agg(**{"WTD OOS Units": ("Qty", "sum"), "OOS Occurrences": ("Product", "size"), "Locations Affected": ("Account", "nunique")}).reset_index()
        wtd_item["% of WTD OOS Units"] = wtd_item["WTD OOS Units"] / wtd_os_units if wtd_os_units else 0
        wtd_item = wtd_item.sort_values(["WTD OOS Units", "OOS Occurrences"], ascending=[False, False]).reset_index(drop=True)
        wtd_item.insert(0, "Rank", range(1, len(wtd_item) + 1))
    else:
        wtd_item = pd.DataFrame(columns=["Rank", "Product", "WTD OOS Units", "OOS Occurrences", "Locations Affected", "% of WTD OOS Units"])

    daily_detail = df.copy()
    if len(daily_detail): daily_detail.insert(0, "Audit Date", str(report_date))

    summary = pd.DataFrame([
        ["Audit Date", str(report_date)],
        ["Workweek", f"Monday {monday:%m/%d/%Y} through {report_date:%m/%d/%Y}"],
        ["Daily Total Units Ordered", total_ordered],
        ["Daily OOS Units", os_units],
        ["Daily OOS %", os_pct],
        ["Daily False OOS Units", false_units],
        ["Daily False OOS Exceptions", len(false_df)],
        ["WTD Total Units Ordered", wtd_ordered],
        ["WTD OOS Units", wtd_os_units],
        ["WTD OOS %", wtd_os_units / wtd_ordered if wtd_ordered else 0],
        ["WTD False OOS Units", int(wtd_false.Qty.sum()) if len(wtd_false) else 0],
        ["WTD False OOS Exceptions", len(wtd_false)],
    ], columns=["Metric", "Value"])

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Executive Summary", index=False, startrow=2)
        daily_loc.to_excel(writer, sheet_name="OOS by Location", index=False)
        item_qty.to_excel(writer, sheet_name="Item Qty", index=False)
        item_freq.to_excel(writer, sheet_name="Item Frequency", index=False)
        repick.to_excel(writer, sheet_name="False OOS Repick", index=False)
        daily_detail.to_excel(writer, sheet_name="Detailed Audit", index=False)
        wtd_loc.to_excel(writer, sheet_name="WTD by Location", index=False)
        wtd_item.to_excel(writer, sheet_name="WTD by Item", index=False)
        history.to_excel(writer, sheet_name="Audit History", index=False)

        wb = writer.book
        dark = "2B2B2B"; light = "F2F2F2"; accent = "D9EAF7"; white = "FFFFFF"
        thin = Side(style="thin", color="B7B7B7")
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else None
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor=dark); cell.font = Font(color=white, bold=True); cell.alignment = Alignment(vertical="center")
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = Border(bottom=thin)
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            for col in range(1, ws.max_column + 1):
                vals = [str(ws.cell(r, col).value or "") for r in range(1, min(ws.max_row, 200) + 1)]
                width = min(max(max((len(v) for v in vals), default=8) + 2, 10), 38)
                ws.column_dimensions[get_column_letter(col)].width = width
            # Percentage formats based on header names
            headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
            for name, col in headers.items():
                if name and "%" in str(name):
                    for r in range(2, ws.max_row + 1): ws.cell(r, col).number_format = "0.00%"

        # Executive Summary has a title above the table and needs special styling.
        ws = wb["Executive Summary"]
        ws.auto_filter.ref = f"A3:B{ws.max_row}"
        ws.freeze_panes = "A4"
        ws["A1"] = "OOS AUDIT — EXECUTIVE SUMMARY"
        ws.merge_cells("A1:B1"); ws["A1"].font = Font(size=16, bold=True, color=white); ws["A1"].fill = PatternFill("solid", fgColor=dark); ws["A1"].alignment = Alignment(horizontal="center")
        for c in ws[3]: c.fill = PatternFill("solid", fgColor=dark); c.font = Font(color=white, bold=True)
        for r in range(4, ws.max_row + 1):
            if "%" in str(ws.cell(r,1).value): ws.cell(r,2).number_format = "0.00%"
            if r % 2 == 0:
                for c in ws[r]: c.fill = PatternFill("solid", fgColor=light)
        ws.column_dimensions["A"].width = 31; ws.column_dimensions["B"].width = 28

        # Make repick sheet print-friendly and visually grouped by location.
        rp = wb["False OOS Repick"]
        rp.page_setup.orientation = "portrait"; rp.page_setup.fitToWidth = 1; rp.sheet_properties.pageSetUpPr.fitToPage = True
        rp.print_title_rows = "1:1"

    out.seek(0)
    return out.getvalue()


st.title("📦 OOS Auditor")
st.caption("Phone-first pick-sheet audit for Android, iPhone and desktop.")
with st.expander("How it works", expanded=False):
    st.write("Column A of Verified Outs = verified OOS. Column B is notes only. An OS on a pick sheet that is not in Column A becomes a picking exception.")

report_date=st.date_input("Audit date",value=date.today())
pick=st.file_uploader("1️⃣ Upload Pick Sheets PDF",type=["pdf"])
verified_file=st.file_uploader("2️⃣ Upload Verified Outs",type=["xlsx","xls","csv","pdf","png","jpg","jpeg"])
prior_history=st.file_uploader("Optional: Upload prior OOS Audit History CSV for Week-to-Date",type=["csv"],help="Upload the history CSV downloaded from the prior audit. The app combines it with today's audit and reports Monday through the selected audit date.")

with st.expander("Manager settings"):
    threshold=st.slider("Product match threshold",.70,1.0,.86,.01)
    model=st.selectbox("Reading model",["gpt-6.1-sol","gpt-6-luna"],index=0)

api_key=os.getenv("OPENAI_API_KEY","")
if not api_key:
    try: api_key=st.secrets["OPENAI_API_KEY"]
    except Exception: api_key=""
if not api_key: st.error("Server setup required: OPENAI_API_KEY has not been configured. Employees should never enter the API key.")

if st.button("RUN AUDIT",type="primary",disabled=not(pick and verified_file and api_key)):
    client=OpenAI(api_key=api_key)
    with st.status("Reading audit...",expanded=True) as s:
        st.write("Reading Verified Outs — Column A only...")
        try:
            vo=verified_outs(verified_file,client,model)
        except Exception as e:
            s.update(label="Verified Outs could not be read",state="error")
            st.error(f"Verified Outs error: {e}")
            st.stop()
        st.write(f"{len(vo)} verified OOS items found.")
        st.write("Reading pick sheets and handwritten OS marks...")
        try:
            data=pick_data(pick,client,model)
        except Exception as e:
            s.update(label="Pick sheets could not be read",state="error")
            st.error(f"Pick sheet error: {e}")
            st.stop()
        st.session_state.update({"data":data,"vo":vo,"report_date":report_date}); s.update(label="Audit ready",state="complete")

if "data" in st.session_state:
    data=st.session_state["data"]; vo=st.session_state["vo"]; rd=st.session_state.get("report_date",report_date)
    df=classify(data,vo,threshold)
    total_ordered=sum(int(o.get("total_ordered",0) or 0) for o in data.get("orders",[])); os_units=int(df.Qty.sum()) if len(df) else 0
    false_units=int(df.loc[df.Status=="FALSE OOS / PICKING EXCEPTION","Qty"].sum()) if len(df) else 0
    st.subheader("Results"); c1,c2=st.columns(2); c1.metric("Ordered",f"{total_ordered:,}"); c2.metric("OS Units",f"{os_units:,}")
    c3,c4=st.columns(2); c3.metric("OS %",f"{(os_units/total_ordered*100 if total_ordered else 0):.2f}%"); c4.metric("False OOS Units",f"{false_units:,}")
    needs=df[df.Status=="NEEDS REVIEW"] if len(df) else df
    if len(needs): st.warning(f"⚠️ {len(needs)} line(s) need manager review before employee accountability.")
    tab1,tab2,tab3=st.tabs(["Exceptions","All OOS","Verified List"])
    with tab1: st.dataframe(df[df.Status!="VERIFIED OOS"] if len(df) else df,use_container_width=True,hide_index=True)
    with tab2: st.dataframe(df,use_container_width=True,hide_index=True)
    with tab3: st.dataframe(pd.DataFrame({"Verified OOS — Column A":vo}),use_container_width=True,hide_index=True)

    today_hist=current_history(rd,data,df); history=combine_history(prior_history,today_hist)
    monday=rd-timedelta(days=rd.weekday()); wtd=history[(history.Date>=monday)&(history.Date<=rd)].copy(); wtd=wtd[pd.to_datetime(wtd.Date).dt.weekday<5]
    st.subheader("Download Reports")
    st.download_button("⬇️ OOS ANALYSIS PDF",analysis_pdf(data,df),f"OOS_Analysis_{rd}.pdf","application/pdf")
    st.download_button("⬇️ FALSE OOS — REPICK LIST",repick_pdf(df),f"False_OOS_Repick_{rd}.pdf","application/pdf")
    st.download_button("⬇️ WEEK-TO-DATE OOS PDF",wtd_pdf(history,rd),f"OOS_WTD_{monday}_to_{rd}.pdf","application/pdf")
    st.download_button("⬇️ COMPLETE EXCEL WORKBOOK",excel_export(rd,data,df,history),f"OOS_Audit_{rd}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",help="One workbook containing daily summaries, item quantity/frequency, false OOS repick, detailed audit data, WTD analysis and audit history.")
    st.download_button("⬇️ OOS AUDIT HISTORY CSV",history.to_csv(index=False).encode("utf-8"),"OOS_Audit_History.csv","text/csv",help="Save this file and upload it with the next day's audit to keep the Monday-Friday WTD report building.")
    st.caption(f"WTD window: {monday:%A %m/%d} through {rd:%A %m/%d}. Saturday/Sunday records are excluded.")
