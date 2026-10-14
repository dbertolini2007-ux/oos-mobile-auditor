import os, io, re, json, tempfile, base64
from difflib import SequenceMatcher

import streamlit as st
import pandas as pd
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

st.set_page_config(page_title="OOS Auditor", page_icon="📦", layout="centered")

st.markdown("""
<style>
.block-container {max-width:760px;padding-top:1rem;padding-bottom:3rem}
div[data-testid="stFileUploader"] {border:2px dashed #777;border-radius:14px;padding:8px}
.stButton button,.stDownloadButton button {min-height:52px;font-size:18px;border-radius:12px;width:100%}
[data-testid="stMetricValue"] {font-size:28px}
@media(max-width:640px){.block-container{padding-left:.8rem;padding-right:.8rem}h1{font-size:1.8rem!important}}
</style>
""", unsafe_allow_html=True)

def get_api_key():
    # Streamlit Community Cloud secrets first; environment variable second.
    try:
        k = st.secrets.get("OPENAI_API_KEY", "")
        if k:
            return str(k)
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY", "")

def norm(x):
    x=re.sub(r"[^a-z0-9]+"," ",str(x or "").lower())
    return " ".join(x.split())

def similarity(a,b):
    a,b=norm(a),norm(b)
    if not a or not b: return 0.0
    if a in b or b in a: return .95
    return SequenceMatcher(None,a,b).ratio()

def parse_json(text):
    text=(text or "").strip()
    text=re.sub(r"^```(?:json)?\s*","",text)
    text=re.sub(r"\s*```$","",text)
    # tolerate explanatory text around JSON
    starts=[x for x in (text.find("{"),text.find("[")) if x >= 0]
    if starts:
        text=text[min(starts):]
    return json.loads(text)

def upload_for_model(client, uploaded):
    suffix=os.path.splitext(uploaded.name)[1] or ".bin"
    with tempfile.NamedTemporaryFile(delete=False,suffix=suffix) as f:
        f.write(uploaded.getvalue()); p=f.name
    try:
        with open(p,"rb") as fh:
            return client.files.create(file=fh,purpose="user_data").id
    finally:
        try: os.unlink(p)
        except: pass

def image_to_data_url(uploaded):
    mime = uploaded.type or "image/jpeg"
    b64=base64.b64encode(uploaded.getvalue()).decode("ascii")
    return f"data:{mime};base64,{b64}"

def verified_outs(uploaded, client, model):
    n=uploaded.name.lower()
    if n.endswith((".xlsx",".xls")):
        df=pd.read_excel(io.BytesIO(uploaded.getvalue()),header=None)
        return [str(v).strip() for v in df.iloc[:,0].dropna() if str(v).strip()]
    if n.endswith(".csv"):
        df=pd.read_csv(io.BytesIO(uploaded.getvalue()),header=None)
        return [str(v).strip() for v in df.iloc[:,0].dropna() if str(v).strip()]

    prompt=('Read this Verified Outs sheet. Anything in the LEFT COLUMN / Column A is a verified OOS item. '
            'Ignore Column B and every other column for determining OOS. Return JSON only: '
            '{"verified_outs":["item 1","item 2"]}')

    # Photos/camera images go directly as vision input, avoiding mobile file-upload quirks.
    if n.endswith((".png",".jpg",".jpeg",".webp")):
        r=client.responses.create(
            model=model,
            input=[{"role":"user","content":[
                {"type":"input_image","image_url":image_to_data_url(uploaded),"detail":"high"},
                {"type":"input_text","text":prompt}
            ]}]
        )
    else:
        fid=upload_for_model(client,uploaded)
        r=client.responses.create(
            model=model,
            input=[{"role":"user","content":[
                {"type":"input_file","file_id":fid},
                {"type":"input_text","text":prompt}
            ]}]
        )
    return [str(x).strip() for x in parse_json(r.output_text).get("verified_outs",[]) if str(x).strip()]

def pick_data(uploaded, client, model):
    fid=upload_for_model(client,uploaded)
    prompt="""Read these vending/micro-market pick sheets carefully.
Return JSON only:
{"orders":[{"account":"name","picker":"name if visible","pages":[1],
"total_ordered":123,
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
    r=client.responses.create(
        model=model,
        input=[{"role":"user","content":[{"type":"input_file","file_id":fid},{"type":"input_text","text":prompt}]}]
    )
    return parse_json(r.output_text)

def classify(data, verified, threshold):
    rows=[]
    for order in data.get("orders",[]):
        for x in order.get("os_lines",[]):
            prod=x.get("product","")
            candidates=[(v,similarity(prod,v)) for v in verified]
            match,score=max(candidates,key=lambda z:z[1]) if candidates else ("",0)
            conf=float(x.get("confidence",0) or 0)
            status="NEEDS REVIEW" if conf < .85 else ("VERIFIED OOS" if score >= threshold else "FALSE OOS / PICKING EXCEPTION")
            rows.append({"Account":order.get("account","Unknown"),"Picker":order.get("picker",""),
                         "Product":prod,"Qty":int(x.get("ordered_qty",0) or 0),"Page":x.get("page",""),
                         "Confidence":round(conf,2),"Closest Verified":match,"Match":round(score,2),"Status":status})
    return pd.DataFrame(rows)

def make_pdf(title, subtitle, headers, rows, widths):
    b=io.BytesIO(); styles=getSampleStyleSheet()
    doc=SimpleDocTemplate(b,pagesize=landscape(letter),leftMargin=24,rightMargin=24,topMargin=24,bottomMargin=24)
    story=[Paragraph(title,styles["Title"]),Paragraph(subtitle,styles["BodyText"]),Spacer(1,10)]
    t=Table([headers]+rows,colWidths=widths,repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#252525")),
                           ("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                           ("GRID",(0,0),(-1,-1),.35,colors.grey),("FONTSIZE",(0,0),(-1,-1),7.5),
                           ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("TOPPADDING",(0,0),(-1,-1),4),
                           ("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    story.append(t); doc.build(story); return b.getvalue()

def audit_pdf(data):
    rows=[]; to=tos=0
    for o in data.get("orders",[]):
        ordered=int(o.get("total_ordered",0) or 0)
        osq=sum(int(x.get("ordered_qty",0) or 0) for x in o.get("os_lines",[]))
        to+=ordered; tos+=osq
        rows.append([o.get("account","Unknown"),ordered,osq,f"{(osq/ordered*100 if ordered else 0):.2f}%"])
    rows.sort(key=lambda r:float(r[3][:-1]),reverse=True)
    rows.append(["TOTAL",to,tos,f"{(tos/to*100 if to else 0):.2f}%"])
    return make_pdf("Out-of-Stock Audit Report","Only pick-sheet lines explicitly marked OS are counted.",
                    ["Account","Ordered Units","OS Units","OS %"],rows,[3.6*inch,1.2*inch,1*inch,1*inch])

def exception_pdf(df):
    f=df[df.Status=="FALSE OOS / PICKING EXCEPTION"] if len(df) else df
    rows=[[r.Account,r.Picker,r.Product,r.Qty,r.Page,r["Closest Verified"],f'{r.Match:.2f}'] for _,r in f.iterrows()]
    if not rows: rows=[["","","No false OOS exceptions found","","","",""]]
    return make_pdf("False OOS / Picking Exception Report",
                    "Exception = marked OS on pick sheet but NOT present in Column A of Verified Outs.",
                    ["Account","Picker","Product","Qty","Page","Closest Verified Item","Match"],rows,
                    [1.4*inch,1*inch,2.3*inch,.45*inch,.45*inch,2*inch,.55*inch])

st.title("📦 OOS Auditor")
st.caption("Phone-first pick-sheet audit for Android, iPhone and desktop.")

with st.expander("How it works"):
    st.write("Column A of Verified Outs = verified OOS. Column B is notes only. "
             "An OS on a pick sheet that is not in Column A becomes a picking exception.")

pick=st.file_uploader("1️⃣ Upload Pick Sheets PDF",type=["pdf"],key="pick_pdf")

st.subheader("2️⃣ Verified Outs")
st.caption("Use either option below. Column A / the left column is the verified OOS list.")

camera=st.camera_input("📷 Take Photo of Verified Outs",key="verified_camera")
verified_upload=st.file_uploader("📁 Or Upload Verified Outs",
    type=["xlsx","xls","csv","pdf","png","jpg","jpeg","webp"],key="verified_file")

verified_file = camera if camera is not None else verified_upload
if camera is not None:
    st.success("Verified Outs photo captured.")
elif verified_upload is not None:
    st.success(f"Verified Outs selected: {verified_upload.name}")

with st.expander("Manager settings"):
    threshold=st.slider("Product match threshold",.70,1.0,.86,.01)
    model=st.selectbox("Reading model",["gpt-6.1-sol","gpt-6-luna"],index=0)

api_key=get_api_key()
if not api_key:
    st.error("Server setup required: OPENAI_API_KEY is missing from Streamlit Secrets.")

ready = pick is not None and verified_file is not None and bool(api_key)
if st.button("RUN AUDIT",type="primary",disabled=not ready):
    client=OpenAI(api_key=api_key)
    try:
        with st.status("Reading today's audit...",expanded=True) as s:
            st.write("Reading Verified Outs — Column A only...")
            vo=verified_outs(verified_file,client,model)
            st.write(f"{len(vo)} verified OOS items found.")
            st.write("Reading pick sheets and handwritten OS marks...")
            data=pick_data(pick,client,model)
            st.session_state["data"]=data; st.session_state["vo"]=vo
            s.update(label="Audit ready",state="complete")
    except Exception as e:
        st.error("The audit could not be completed.")
        st.exception(e)

if "data" in st.session_state:
    data=st.session_state["data"]; vo=st.session_state["vo"]
    df=classify(data,vo,threshold)
    total_ordered=sum(int(o.get("total_ordered",0) or 0) for o in data.get("orders",[]))
    os_units=int(df.Qty.sum()) if len(df) else 0
    false_units=int(df.loc[df.Status=="FALSE OOS / PICKING EXCEPTION","Qty"].sum()) if len(df) else 0
    st.subheader("Results")
    c1,c2=st.columns(2); c1.metric("Ordered",f"{total_ordered:,}"); c2.metric("OS Units",f"{os_units:,}")
    c3,c4=st.columns(2); c3.metric("OS %",f"{(os_units/total_ordered*100 if total_ordered else 0):.2f}%"); c4.metric("False OOS Units",f"{false_units:,}")
    needs=df[df.Status=="NEEDS REVIEW"] if len(df) else df
    if len(needs): st.warning(f"⚠️ {len(needs)} line(s) need manager review before employee accountability.")
    t1,t2,t3=st.tabs(["Exceptions","All OS","Verified List"])
    with t1: st.dataframe(df[df.Status!="VERIFIED OOS"] if len(df) else df,use_container_width=True,hide_index=True)
    with t2: st.dataframe(df,use_container_width=True,hide_index=True)
    with t3: st.dataframe(pd.DataFrame({"Verified OOS — Column A":vo}),use_container_width=True,hide_index=True)
    st.subheader("Download Reports")
    st.download_button("⬇️ OOS AUDIT PDF",audit_pdf(data),"OOS_Audit_Report.pdf","application/pdf")
    st.download_button("⬇️ FALSE OOS REPORT",exception_pdf(df),"False_OOS_Picking_Exceptions.pdf","application/pdf")
