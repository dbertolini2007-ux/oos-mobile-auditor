import os, io, re, json, tempfile, base64, hashlib, urllib.request, urllib.error
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
try:
    from PIL import Image
except Exception:
    Image = None

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


def _optimized_image_bytes(raw, max_side=1800, quality=82):
    """Shrink large phone photos before API upload. Keeps enough detail for pick-sheet handwriting."""
    if Image is None:
        return raw, "image/jpeg"
    try:
        im = Image.open(io.BytesIO(raw))
        im = im.convert("RGB")
        if max(im.size) > max_side:
            scale = max_side / max(im.size)
            im = im.resize((max(1, int(im.width*scale)), max(1, int(im.height*scale))))
        out = io.BytesIO(); im.save(out, format="JPEG", quality=quality, optimize=True)
        return out.getvalue(), "image/jpeg"
    except Exception:
        return raw, "image/jpeg"

def model_attachment(uploaded, client):
    """Images go to Responses as input_image; PDFs/docs use Files + input_file."""
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        raw, mime = _optimized_image_bytes(raw)
        b64 = base64.b64encode(raw).decode("ascii")
        return {"type":"input_image", "image_url":f"data:{mime};base64,{b64}", "detail":"high"}
    fid = upload_for_model(client, uploaded)
    return {"type":"input_file", "file_id":fid}

def verified_outs(uploaded, client, model):
    n = uploaded.name.lower()
    if n.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(uploaded.getvalue()), header=None)
        return [str(v).strip() for v in df.iloc[:, 0].dropna() if str(v).strip()]
    if n.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(uploaded.getvalue()), header=None)
        return [str(v).strip() for v in df.iloc[:, 0].dropna() if str(v).strip()]
    attachment = model_attachment(uploaded, client)
    r = client.responses.create(model=model, input=[{"role":"user","content":[
        attachment,
        {"type":"input_text","text":'Read this Verified Outs sheet. Anything in the LEFT COLUMN / Column A is a verified OOS item. Ignore Column B and all other columns. Return JSON only: {"verified_outs":["item 1","item 2"]}'}]}])
    return parse_json(r.output_text).get("verified_outs", [])


def pick_data(uploaded_files, client, model):
    # Accept one PDF, one photo, or multiple photos/PDFs from a phone camera/gallery.
    if not isinstance(uploaded_files, (list, tuple)):
        uploaded_files = [uploaded_files]
    content = [model_attachment(uploaded, client) for uploaded in uploaded_files]
    prompt = """Read all attached vending/micro-market route/pick sheets carefully. Attachments may be PDFs or phone photos, and multiple photos may be consecutive pages of the same pick list. Treat all attachments as one audit batch and consolidate pages belonging to the same order/account.
Extract the ROUTE DATE / PICK-SHEET DATE printed on the sheets. Do not use today's date. If the same route date appears on the batch, return it once as route_date in YYYY-MM-DD format. If dates conflict, use the date printed on the majority of the route sheets and lower date_confidence.
Only lines explicitly marked OS count as out of stock. Checkmarks, blanks, partial fills, quantities written for fills, and other notes are not OS.
Do not double-count cumulative printed totals on continuation pages.
Also return every distinct product/SKU printed on each location's pick sheet in all_products. This is required to calculate product availability, even when the product is not marked OS.
Return JSON only:
{"route_date":"YYYY-MM-DD","date_confidence":0.98,"orders":[{"account":"name","picker":"name if visible","pages":[1],"total_ordered":123,
"all_products":["product 1","product 2"],
"os_lines":[{"product":"full product name","ordered_qty":5,"page":1,"confidence":0.98}]}]}
Use the ordered quantity for each OS line, not a handwritten filled quantity. If handwriting or product text is uncertain, lower confidence rather than guessing."""
    content.append({"type":"input_text","text":prompt})
    r = client.responses.create(model=model, input=[{"role":"user","content":content}])
    return parse_json(r.output_text)


def _upload_fingerprint(files):
    if not isinstance(files, (list, tuple)): files=[files]
    h=hashlib.sha256()
    for f in files:
        h.update(str(f.name).encode()); h.update(f.getvalue())
    return h.hexdigest()


def cached_verified_outs(uploaded, client, model):
    key=("verified", model, _upload_fingerprint(uploaded))
    cache=st.session_state.setdefault("ai_result_cache", {})
    if key not in cache: cache[key]=verified_outs(uploaded, client, model)
    return cache[key]

def cached_pick_data(uploaded_files, client, model):
    key=("pick", model, _upload_fingerprint(uploaded_files))
    cache=st.session_state.setdefault("ai_result_cache", {})
    if key not in cache: cache[key]=pick_data(uploaded_files, client, model)
    return cache[key]

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


def _route_date(data, fallback):
    raw = str(data.get("route_date", "") or "").strip()
    try:
        return pd.to_datetime(raw, errors="raise").date()
    except Exception:
        return fallback


def _product_key(x):
    return norm(x)


def analysis_pdf(data, df, report_date=None):
    """One-page executive OOS report matching the approved example layout."""
    b = io.BytesIO(); styles = getSampleStyleSheet()
    route_date = _route_date(data, report_date or date.today())
    title_style = ParagraphStyle("report_title", parent=styles["Title"], fontSize=20, leading=22, alignment=TA_CENTER, spaceAfter=2)
    date_style = ParagraphStyle("route_date", parent=styles["BodyText"], fontSize=10, leading=12, alignment=TA_CENTER, textColor=DARK, spaceAfter=1)
    sub_style = ParagraphStyle("report_sub", parent=styles["BodyText"], fontSize=8, leading=10, textColor=colors.HexColor("#555555"), alignment=TA_CENTER, spaceAfter=6)
    h2 = ParagraphStyle("h2exec", parent=styles["Heading2"], fontSize=12.5, leading=14, textColor=DARK, spaceBefore=5, spaceAfter=4)
    doc = SimpleDocTemplate(b, pagesize=letter, leftMargin=28, rightMargin=28, topMargin=18, bottomMargin=18)
    story = [Paragraph("Out-of-Stock Analysis", title_style),
             Paragraph(f"Route Date: {route_date:%B %d, %Y}", date_style),
             Paragraph("Only pick-sheet lines explicitly marked OS are counted.", sub_style)]

    # Consolidate continuation pages/order blocks by location.
    loc_map = {}
    overall_products = {}
    for o in data.get("orders", []):
        account = str(o.get("account", "Unknown") or "Unknown").strip()
        ordered = int(o.get("total_ordered", 0) or 0)
        os_lines = o.get("os_lines", []) or []
        all_products = o.get("all_products", []) or []
        rec = loc_map.setdefault(account, {"ordered":0, "os_units":0, "products":{}, "oos_products":{}})
        rec["ordered"] += ordered
        rec["os_units"] += sum(int(x.get("ordered_qty", 0) or 0) for x in os_lines)
        for prod in all_products:
            k = _product_key(prod)
            if k:
                rec["products"][k] = str(prod)
                overall_products[k] = str(prod)
        for x in os_lines:
            prod = str(x.get("product", "") or "").strip(); k = _product_key(prod)
            if k:
                rec["oos_products"][k] = prod
                # Older extraction results may not contain all_products; keep OOS products represented.
                rec["products"].setdefault(k, prod)
                overall_products.setdefault(k, prod)

    total_ordered = sum(v["ordered"] for v in loc_map.values())
    os_units = int(df.Qty.sum()) if len(df) else 0
    false_df = df[df.Status == "FALSE OOS / PICKING EXCEPTION"] if len(df) else df
    false_units = int(false_df.Qty.sum()) if len(false_df) else 0
    unique_products = len(overall_products)
    overall_oos_products = {_product_key(p) for p in (df.Product.tolist() if len(df) else []) if _product_key(p)}
    unique_oos = len(overall_oos_products)
    in_stock_pct = ((unique_products - unique_oos) / unique_products * 100) if unique_products else 0
    prod_oos_pct = (unique_oos / unique_products * 100) if unique_products else 0

    def kpi_table(headers, vals):
        t=Table([headers, vals], colWidths=[1.79*inch]*4, rowHeights=[.29*inch,.34*inch])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,0),7.2),
            ("ALIGN",(0,0),(-1,-1),"CENTER"),("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),
            ("FONTSIZE",(0,1),(-1,1),12.5),("BOX",(0,0),(-1,-1),.4,colors.grey),
            ("INNERGRID",(0,0),(-1,-1),.25,colors.grey),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
        story.append(t); story.append(Spacer(1,4))

    unit_pct = os_units/total_ordered*100 if total_ordered else 0
    kpi_table(["TOTAL UNITS ORDERED","OOS UNITS","OOS %","FALSE OOS UNITS"],
              [f"{total_ordered:,}",f"{os_units:,}",f"{unit_pct:.2f}%",f"{false_units:,}"])
    kpi_table(["UNIQUE PRODUCTS","UNIQUE PRODUCTS OOS","PRODUCT IN-STOCK %","PRODUCT OOS %"],
              [f"{unique_products:,}",f"{unique_oos:,}",f"{in_stock_pct:.2f}%",f"{prod_oos_pct:.2f}%"])

    story.append(Paragraph("Product Availability by Location", h2))
    avail=[]
    for account,v in loc_map.items():
        up=len(v["products"]); uo=len(v["oos_products"])
        inst=((up-uo)/up*100) if up else 0; oop=(uo/up*100) if up else 0
        avail.append([account, f"{up:,}", f"{uo:,}", f"{inst:.2f}%", f"{oop:.2f}%"])
    avail.sort(key=lambda r: float(r[4].replace("%","")), reverse=True)
    avail.append(["TOTAL", f"{unique_products:,}", f"{unique_oos:,}", f"{in_stock_pct:.2f}%", f"{prod_oos_pct:.2f}%"])
    t=Table([["Location","Unique Products","Unique OOS","In-Stock %","OOS %"]]+avail,
            colWidths=[3.05*inch,1.08*inch,.95*inch,.95*inch,.85*inch], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),6.8),("GRID",(0,0),(-1,-1),.3,colors.HexColor("#AAAAAA")),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[colors.white,LIGHT]),("BACKGROUND",(0,-1),(-1,-1),MID),("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
        ("TOPPADDING",(0,0),(-1,-1),3.2),("BOTTOMPADDING",(0,0),(-1,-1),3.2),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    story.append(t)

    story.append(Paragraph("Out-of-Stock Percentage by Location", h2))
    loc=[]
    for account,v in loc_map.items():
        pct=v["os_units"]/v["ordered"]*100 if v["ordered"] else 0
        loc.append([account,f"{v['ordered']:,}",f"{v['os_units']:,}",f"{pct:.2f}%"])
    loc.sort(key=lambda r: float(r[3].replace("%","")), reverse=True)
    loc.append(["TOTAL",f"{total_ordered:,}",f"{os_units:,}",f"{unit_pct:.2f}%"])
    t2=Table([["Location","Total Units Ordered","Units on OOS Lines","OOS %"]]+loc,
             colWidths=[3.25*inch,1.45*inch,1.45*inch,.85*inch], repeatRows=1)
    t2.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),6.8),("GRID",(0,0),(-1,-1),.3,colors.HexColor("#AAAAAA")),
        ("ROWBACKGROUNDS",(0,1),(-1,-2),[colors.white,LIGHT]),("BACKGROUND",(0,-1),(-1,-1),MID),("FONTNAME",(0,-1),(-1,-1),"Helvetica-Bold"),
        ("TOPPADDING",(0,0),(-1,-1),3.2),("BOTTOMPADDING",(0,0),(-1,-1),3.2),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    story.append(t2)
    doc.build(story)
    return b.getvalue()

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


def _secret(name, default=""):
    try: return str(st.secrets.get(name, default) or default)
    except Exception: return default

def _github_history_config():
    return (_secret("WTD_GITHUB_TOKEN"), _secret("WTD_GITHUB_REPO"), _secret("WTD_GITHUB_BRANCH","main"), _secret("WTD_GITHUB_PATH","data/oos_wtd_history.csv"))

def _github_request(url, token, method="GET", payload=None):
    headers={"Authorization":f"Bearer {token}","Accept":"application/vnd.github+json","X-GitHub-Api-Version":"2022-11-28","User-Agent":"OOS-Mobile-Auditor"}
    data=json.dumps(payload).encode() if payload is not None else None
    req=urllib.request.Request(url,data=data,headers=headers,method=method)
    with urllib.request.urlopen(req,timeout=20) as resp:
        return json.loads(resp.read().decode() or "{}")

def load_persistent_history():
    """Load centralized WTD history from a PRIVATE GitHub data repo when configured."""
    token, repo, branch, path=_github_history_config()
    if not token or not repo: return pd.DataFrame(), False
    url=f"https://api.github.com/repos/{repo}/contents/{path}?ref={branch}"
    try:
        obj=_github_request(url,token)
        raw=base64.b64decode(obj.get("content","").replace("\\n",""))
        return pd.read_csv(io.BytesIO(raw)), True
    except urllib.error.HTTPError as e:
        if e.code==404: return pd.DataFrame(), True
        raise

def save_persistent_history(history):
    token, repo, branch, path=_github_history_config()
    if not token or not repo: return False
    url=f"https://api.github.com/repos/{repo}/contents/{path}"
    sha=None
    try: sha=_github_request(url+f"?ref={branch}",token).get("sha")
    except urllib.error.HTTPError as e:
        if e.code!=404: raise
    csv_bytes=history.to_csv(index=False).encode()
    payload={"message":"Update OOS WTD history","content":base64.b64encode(csv_bytes).decode(),"branch":branch}
    if sha: payload["sha"]=sha
    _github_request(url,token,"PUT",payload)
    return True

def merge_replace_day(history, today_df, report_date):
    """Replace an existing route date instead of double-counting a rerun."""
    parts=[]
    if history is not None and len(history):
        h=history.copy(); h["Date"]=pd.to_datetime(h["Date"],errors="coerce").dt.date
        h=h[h["Date"] != report_date]
        parts.append(h)
    t=today_df.copy(); t["Date"]=report_date; parts.append(t)
    return pd.concat(parts,ignore_index=True) if parts else t

def current_history(report_date, data, df):
    ordered_map={str(o.get("account","Unknown")):int(o.get("total_ordered",0) or 0) for o in data.get("orders",[])}
    unique_map={str(o.get("account","Unknown")):len(set(norm(x) for x in o.get("all_products",[]) if str(x).strip())) for o in data.get("orders",[])}
    rows=[]
    if len(df):
        for r in df.itertuples(index=False):
            rows.append({"Date":str(report_date),"Account":r.Account,"Picker":r.Picker,"Product":r.Product,"Qty":int(r.Qty),"Status":r.Status,"OrderedUnits":ordered_map.get(str(r.Account),0),"UniqueProducts":unique_map.get(str(r.Account),0)})
    # Preserve locations with zero OOS so WTD ordered-unit and product-availability totals stay correct.
    present=set(df.Account.astype(str)) if len(df) else set()
    for account,ordered in ordered_map.items():
        if account not in present:
            rows.append({"Date":str(report_date),"Account":account,"Picker":"","Product":"","Qty":0,"Status":"NO OOS","OrderedUnits":ordered,"UniqueProducts":unique_map.get(account,0)})
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


def filter_wtd(history, report_date):
    """Return Monday-Friday WTD rows with dates normalized to pandas Timestamps.
    This avoids comparing Python date objects with datetime64 columns after Streamlit reruns.
    """
    h = history.copy()
    if "Date" not in h.columns:
        return h
    h["Date"] = pd.to_datetime(h["Date"], errors="coerce").dt.normalize()
    h = h[h["Date"].notna()].copy()
    end = pd.Timestamp(report_date).normalize()
    monday = end - pd.Timedelta(days=end.weekday())
    return h[(h["Date"] >= monday) & (h["Date"] <= end) & (h["Date"].dt.weekday < 5)].copy()


def wtd_pdf(history, report_date):
    """WTD PDF: page 1 mirrors the normal OOS executive report; later pages retain weekly detail."""
    monday=report_date-timedelta(days=report_date.weekday())
    h=filter_wtd(history, report_date)
    b=io.BytesIO(); doc=SimpleDocTemplate(b,pagesize=landscape(letter),rightMargin=18,leftMargin=18,topMargin=16,bottomMargin=16)
    styles=getSampleStyleSheet(); title=ParagraphStyle("wt",parent=styles["Title"],fontSize=20,leading=22,alignment=TA_CENTER,spaceAfter=2)
    sub=ParagraphStyle("ws",parent=styles["BodyText"],fontSize=9,leading=11,alignment=TA_CENTER,spaceAfter=7)
    h2=ParagraphStyle("wh",parent=styles["Heading2"],fontSize=13,leading=15,spaceBefore=5,spaceAfter=5)
    story=[Paragraph("Week-to-Date Out-of-Stock Analysis",title),Paragraph(f"Report Date Range: {monday:%B %d, %Y} – {report_date:%B %d, %Y} (Monday–Friday)",sub)]
    ord_units=int(h[["Date","Account","OrderedUnits"]].drop_duplicates(["Date","Account"])["OrderedUnits"].sum()) if len(h) else 0
    os_h=h[(h.Product.fillna("").astype(str).str.strip()!="") & (h.Status!="NO OOS")]
    os_units=int(os_h.Qty.sum()) if len(os_h) else 0
    false_h=os_h[os_h.Status=="FALSE OOS / PICKING EXCEPTION"]
    # Product KPIs use distinct location/product pairs across the selected week.
    hp=h[h.Product.fillna("").astype(str).str.strip()!=""].copy() if len(h) else h.copy()
    unique_products=int(hp.Product.nunique()) if len(hp) else 0
    unique_oos=int(os_h.Product.nunique()) if len(os_h) else 0
    instock=(unique_products-unique_oos)/unique_products if unique_products else 0
    product_oos=unique_oos/unique_products if unique_products else 0
    k1=Table([["TOTAL UNITS ORDERED","OOS UNITS","OOS %","FALSE OOS UNITS"],[f"{ord_units:,}",f"{os_units:,}",f"{(os_units/ord_units*100 if ord_units else 0):.2f}%",f"{int(false_h.Qty.sum()) if len(false_h) else 0:,}"]],colWidths=[2.35*inch]*4)
    k1.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),("FONTSIZE",(0,0),(-1,0),8),("FONTSIZE",(0,1),(-1,1),15),("GRID",(0,0),(-1,-1),.4,colors.grey),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)])); story += [k1,Spacer(1,5)]
    k2=Table([["UNIQUE PRODUCTS","UNIQUE PRODUCTS OOS","PRODUCT IN-STOCK %","PRODUCT OOS %"],[f"{unique_products:,}",f"{unique_oos:,}",f"{instock*100:.2f}%",f"{product_oos*100:.2f}%"]],colWidths=[2.35*inch]*4)
    k2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),DARK),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),("FONTSIZE",(0,0),(-1,0),8),("FONTSIZE",(0,1),(-1,1),15),("GRID",(0,0),(-1,-1),.4,colors.grey),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)])); story += [k2,Spacer(1,6)]
    # Availability by location. History has OS product rows; unique product totals are stored per location/day below.
    story.append(Paragraph("Product Availability by Location",h2))
    avail=[]
    for account,g in h.groupby("Account") if len(h) else []:
        # Sum daily unique-product counts if present; for WTD we use max product assortment seen at that location during week.
        up=int(pd.to_numeric(g.get("UniqueProducts",pd.Series([0])),errors="coerce").fillna(0).max())
        go=os_h[os_h.Account==account] if len(os_h) else os_h
        uo=int(go.Product.nunique()) if len(go) else 0
        avail.append([account,up,uo,f"{((up-uo)/up*100 if up else 0):.2f}%",f"{(uo/up*100 if up else 0):.2f}%"])
    avail.sort(key=lambda r:float(r[4][:-1]),reverse=True)
    add_table(story,["Location","Unique Products","Unique OOS","In-Stock %","OOS %"],avail,[3.7*inch,1.4*inch,1.2*inch,1.2*inch,1.0*inch],7.5)
    story += [Spacer(1,5),Paragraph("Out-of-Stock Percentage by Location",h2)]
    locrows=[]
    for account,g in h.groupby("Account") if len(h) else []:
        ordered=int(g[["Date","Account","OrderedUnits"]].drop_duplicates(["Date","Account"])["OrderedUnits"].sum())
        og=os_h[os_h.Account==account] if len(os_h) else os_h; oq=int(og.Qty.sum()) if len(og) else 0
        locrows.append([account,ordered,oq,f"{(oq/ordered*100 if ordered else 0):.2f}%"])
    locrows.sort(key=lambda r:float(r[3][:-1]),reverse=True); locrows.append(["TOTAL",ord_units,os_units,f"{(os_units/ord_units*100 if ord_units else 0):.2f}%"])
    add_table(story,["Location","Total Units Ordered","Units on OOS Lines","OOS %"],locrows,[4.0*inch,1.8*inch,1.8*inch,1.2*inch],7.5)
    # Existing weekly detail begins page 2.
    story += [PageBreak(),Paragraph("Week-to-Date Item Detail",title),Paragraph(f"{monday:%m/%d/%Y} through {report_date:%m/%d/%Y}",sub),Paragraph("1. OOS by Individual Item Quantity",h2)]
    if len(os_h):
        q=os_h.groupby("Product").agg(OS_Units=("Qty","sum"),Locations=("Account","nunique")).reset_index().sort_values("OS_Units",ascending=False)
        qr=[[i,r.Product,int(r.OS_Units),int(r.Locations),f"{(r.OS_Units/os_units*100 if os_units else 0):.1f}%"] for i,r in enumerate(q.itertuples(index=False),1)]; qr.append(["","TOTAL",os_units,"","100.0%"])
    else: qr=[["","No OOS items",0,0,"0.0%"]]
    add_table(story,["Rank","Product","WTD OOS Units","Locations OOS","% of WTD OOS"],qr,[.55*inch,4.4*inch,1.0*inch,1.0*inch,1.0*inch],7.5)
    story += [Spacer(1,16),Paragraph("2. OOS by Item Frequency / Occurrences",h2)]
    if len(os_h):
        f=os_h.groupby("Product").agg(Occurrences=("Product","size"),Locations=("Account","nunique"),OS_Units=("Qty","sum")).reset_index().sort_values(["Occurrences","OS_Units"],ascending=False)
        fr=[[i,r.Product,int(r.Occurrences),int(r.Locations),int(r.OS_Units)] for i,r in enumerate(f.itertuples(index=False),1)]; fr.append(["","TOTAL",len(os_h),"",os_units])
    else: fr=[["","No OOS items",0,0,0]]
    add_table(story,["Rank","Product","WTD OOS Occurrences","Locations OOS","WTD OOS Units"],fr,[.55*inch,4.3*inch,1.2*inch,1.0*inch,1.0*inch],7.5)
    doc.build(story); return b.getvalue()


def excel_export(report_date, data, df, history):
    """Create one polished Excel workbook containing the daily audit and Monday-Friday WTD analysis."""
    monday = report_date - timedelta(days=report_date.weekday())
    h = filter_wtd(history, report_date)

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

# Pick sheets: employees can either take photos directly or select existing pictures/PDFs.
if "camera_pick_photos" not in st.session_state:
    st.session_state.camera_pick_photos = []

st.markdown("### 1️⃣ Pick Sheets")
pick_source = st.radio(
    "Choose how to add pick sheets",
    ["🖼️ Select Pictures / Files", "📷 Take Picture"],
    horizontal=True,
    key="pick_source",
)

# IMPORTANT: The camera widget is created ONLY after the employee explicitly
# selects Take Picture. This prevents the camera from opening automatically.
selected_pick_files = []
if pick_source == "📷 Take Picture":
    camera_photo = st.camera_input("Take a picture of a picklist page", key="pick_camera")
    if camera_photo is not None:
        if st.button("➕ ADD THIS PHOTO", use_container_width=True):
            raw = camera_photo.getvalue()
            # Avoid adding the exact same capture twice.
            if not any(x["data"] == raw for x in st.session_state.camera_pick_photos):
                st.session_state.camera_pick_photos.append({"name": f"picklist_camera_{len(st.session_state.camera_pick_photos)+1}.jpg", "data": raw})
                st.success(f"Photo added. {len(st.session_state.camera_pick_photos)} camera photo(s) ready.")
            else:
                st.info("That photo is already added.")
    if st.session_state.camera_pick_photos:
        st.caption(f"{len(st.session_state.camera_pick_photos)} camera photo(s) added. Take the next page and tap ADD THIS PHOTO again.")
        if st.button("🗑️ CLEAR CAMERA PHOTOS", use_container_width=True):
            st.session_state.camera_pick_photos = []
            st.rerun()
else:
    selected_pick_files = st.file_uploader(
        "Select existing pictures or PDF pick sheets",
        type=["pdf","png","jpg","jpeg","webp"],
        accept_multiple_files=True,
        help="Select one or multiple pictures/PDFs from the phone.",
        key="pick_file_upload",
    )

class MemoryUpload:
    def __init__(self, name, data):
        self.name = name
        self._data = data
    def getvalue(self):
        return self._data

pick = [MemoryUpload(x["name"], x["data"]) for x in st.session_state.camera_pick_photos] + list(selected_pick_files or [])
if pick:
    st.success(f"✅ {len(pick)} pick-sheet page/file(s) ready for audit.")

st.markdown("### 2️⃣ Verified Outs")
verified_file=st.file_uploader("Upload Verified Outs",type=["xlsx","xls","csv","pdf","png","jpg","jpeg"],key="verified_outs_upload")

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
    try:
        with st.status("Reading audit...",expanded=True) as s:
            st.write("Reading Verified Outs — Column A only..."); vo=cached_verified_outs(verified_file,client,model); st.write(f"{len(vo)} verified OOS items found.")
            st.write(f"Reading {len(pick)} pick-sheet file(s)/photo(s) and handwritten OS marks..."); data=cached_pick_data(pick,client,model)
            route_rd = _route_date(data, report_date)
            st.session_state.update({"data":data,"vo":vo,"report_date":route_rd}); s.update(label="Audit ready",state="complete")
    except Exception as e:
        st.error("The audit reader could not process one of the uploads. Try again, or use a clear JPG/PNG photo or PDF.")
        st.caption(f"Technical detail: {type(e).__name__}: {str(e)[:500]}")

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

    today_hist=current_history(rd,data,df)
    # Centralized WTD storage when a private GitHub data repo is configured.
    # Re-running the same route date REPLACES that date so it is never double-counted.
    try:
        persisted, persistent_enabled = load_persistent_history()
        base_hist = persisted if persistent_enabled else st.session_state.get("wtd_history", pd.DataFrame())
        history = merge_replace_day(base_hist, today_hist, rd)
        if persistent_enabled:
            save_persistent_history(history)
        st.session_state.wtd_history = history.copy()
        st.session_state.wtd_storage_mode = "persistent" if persistent_enabled else "session"
    except Exception as e:
        base_hist=st.session_state.get("wtd_history", pd.DataFrame())
        history=merge_replace_day(base_hist,today_hist,rd); st.session_state.wtd_history=history.copy()
        st.session_state.wtd_storage_mode="session"
        st.warning(f"WTD permanent storage is temporarily unavailable; this audit is retained for this session. {type(e).__name__}")
    monday=rd-timedelta(days=rd.weekday()); wtd=filter_wtd(history, rd)
    st.subheader("Download Reports")
    st.download_button("⬇️ OOS ANALYSIS — ONE-PAGE REPORT",analysis_pdf(data,df,rd),f"OOS_Analysis_{rd}.pdf","application/pdf")
    st.download_button("⬇️ FALSE OOS — REPICK LIST",repick_pdf(df),f"False_OOS_Repick_{rd}.pdf","application/pdf")
    st.download_button("⬇️ WEEK-TO-DATE OOS PDF",wtd_pdf(history,rd),f"OOS_WTD_{monday}_to_{rd}.pdf","application/pdf")
    st.download_button("⬇️ COMPLETE EXCEL WORKBOOK",excel_export(rd,data,df,history),f"OOS_Audit_{rd}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",help="One workbook containing daily summaries, item quantity/frequency, false OOS repick, detailed audit data and WTD analysis.")
    st.caption(f"WTD window: {monday:%A %m/%d} through {rd:%A %m/%d}. Saturday/Sunday records are excluded.")
    if st.session_state.get("wtd_storage_mode")=="persistent": st.caption("✅ WTD data saved centrally. Re-running the same route date replaces that day instead of duplicating it.")
    else: st.caption("ℹ️ WTD is currently session-only. Configure private persistent storage secrets to retain data across Streamlit restarts and for all team members.")
