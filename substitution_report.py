"""Optional verified-OOS substitution report. No changes to existing report logic."""
import io
import re
from difflib import SequenceMatcher
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer


def clean(value):
    return re.sub(r"[^A-Z0-9]+", " ", str(value or "").upper()).strip()


def catalog_from_excel(file):
    book = pd.ExcelFile(file)
    frames = []
    for sheet in book.sheet_names:
        frame = pd.read_excel(book, sheet_name=sheet, header=None)
        for i in range(min(12, len(frame))):
            headers = [clean(x).replace(" ", "") for x in frame.iloc[i].tolist()]
            if "RUTHID" not in headers:
                continue
            names = frame.iloc[i].astype(str).tolist()
            data = frame.iloc[i+1:].copy()
            data.columns = names
            mapping = {clean(c).replace(" ", ""): c for c in data.columns}
            id_col = mapping.get("RUTHID")
            desc_col = mapping.get("ITEM") or mapping.get("DESCRIPTION") or mapping.get("PRODUCT")
            if not id_col or not desc_col:
                continue
            result = pd.DataFrame({
                "RuthID": data[id_col].astype(str).str.replace(r"\\.0$", "", regex=True).str.strip(),
                "Item": data[desc_col].fillna("").astype(str).str.strip(),
                "Category": data[mapping["CATGRY"]].fillna("").astype(str) if "CATGRY" in mapping else "",
                "Vendor": data[mapping["VENDOR"]].fillna("").astype(str) if "VENDOR" in mapping else "",
                "UnitsPerCase": pd.to_numeric(data[mapping["UNITS/CASE"]], errors="coerce") if "UNITS/CASE" in mapping else 0,
            })
            frames.append(result)
            break
    if not frames:
        raise ValueError("Catalog must contain RuthID and Item columns.")
    out = pd.concat(frames, ignore_index=True)
    return out[out.RuthID.str.fullmatch(r"\\d+") & out.Item.ne("")].drop_duplicates("RuthID")


def best_match(description, catalog):
    key = clean(description)
    scores = catalog.copy()
    scores["Score"] = scores.Item.map(lambda x: SequenceMatcher(None, key, clean(x)).ratio())
    scores = scores.sort_values("Score", ascending=False)
    return scores.iloc[0] if len(scores) else None


def candidates(item, catalog, excluded):
    orig = best_match(item, catalog)
    if orig is None:
        return None, []
    pool = catalog[~catalog.RuthID.isin(excluded | {orig.RuthID})].copy()
    if not len(pool):
        return orig, []
    orig_words = set(clean(orig.Item).split())
    def score(row):
        words = set(clean(row.Item).split())
        similarity = SequenceMatcher(None, clean(orig.Item), clean(row.Item)).ratio()
        category = 0.22 if clean(row.Category) and clean(row.Category) == clean(orig.Category) else 0
        vendor = 0.12 if clean(row.Vendor) and clean(row.Vendor) == clean(orig.Vendor) else 0
        case = 0.06 if row.UnitsPerCase and row.UnitsPerCase == orig.UnitsPerCase else 0
        shared = 0.10 * len(orig_words & words) / max(1, len(orig_words))
        return similarity + category + vendor + case + shared
    pool["MatchScore"] = pool.apply(score, axis=1)
    return orig, pool.sort_values("MatchScore", ascending=False).head(8).to_dict("records")


def pdf_report(rows, report_date):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = [Paragraph("SUGGESTED SUBSTITUTION PICKLIST — REPORT #4", styles["Title"]),
             Paragraph("Verified actual OOS only | Employee-approved replacement items", styles["Normal"]),
             Paragraph(f"Route date: {escape(str(report_date))}", styles["Normal"]), Spacer(1, 14)]
    if not rows:
        story.append(Paragraph("No substitutions approved.", styles["Normal"]))
    for account in dict.fromkeys(str(r["Account"]) for r in rows):
        group = [r for r in rows if str(r["Account"]) == account]
        story.append(Paragraph(escape(account), styles["Heading2"]))
        cells = [["QTY", "OOS RUTHID / ITEM", "SUB RUTHID / ITEM"]]
        for r in group:
            cells.append([
                str(int(r["Qty"])),
                Paragraph(escape(str(r["OriginalRuthID"])) + " - " + escape(str(r["Product"])), styles["BodyText"]),
                Paragraph(escape(str(r["ReplacementRuthID"])) + " - " + escape(str(r["ReplacementItem"])), styles["BodyText"])
            ])
        table = Table(cells, colWidths=[0.5*inch, 3.0*inch, 3.8*inch], repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#25374C")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("TOPPADDING", (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F1F4F8")]),
        ]))
        story.extend([table, Spacer(1, 14)])
    doc.build(story)
    return buf.getvalue()


def render_report_4(df, rd):
    st.markdown("### Report #4 — Suggested Substitution Picklist (Optional)")
    st.caption("Only VERIFIED OOS rows are eligible. Confirm product identity and replacement suitability before printing.")
    upload = st.file_uploader("Product catalog (ITEMS.xlsx)", type=["xlsx"], key="sub_catalog")
    if upload is None:
        st.info("Upload the product catalog to enable substitution suggestions.")
        return
    try:
        catalog = catalog_from_excel(upload)
    except Exception as e:
        st.error(f"Cannot read catalog: {e}")
        return
    verified = df[df["Status"].eq("VERIFIED OOS")].copy() if len(df) else pd.DataFrame()
    if verified.empty:
        st.info("No confirmed verified OOS lines to substitute.")
        return
    excluded = set()
    for product in verified.Product.unique():
        original = best_match(product, catalog)
        if original is not None:
            excluded.add(str(original.RuthID))
    rows = []
    for n, item in enumerate(verified.itertuples(index=False)):
        orig, choices = candidates(item.Product, catalog, excluded)
        original_id = str(orig.RuthID) if orig is not None else "UNMATCHED"
        # Require explicit employee approval, including for the top-ranked suggestion.
        options = ["NO SUBSTITUTION"] + [str(x["RuthID"]) + " | " + str(x["Item"]) for x in choices]
        st.markdown(f"**{item.Account} — {item.Product} — Qty {int(item.Qty)}**")
        st.caption(f"Original catalog RuthID: {original_id} (verify before approval)")
        selected = st.selectbox("Replacement", options, key=f"sub_pick_{n}")
        approved = st.checkbox("Approve replacement", value=False, key=f"sub_approve_{n}", disabled=selected=="NO SUBSTITUTION")
        if approved and selected != "NO SUBSTITUTION":
            replacement_id = selected.split(" | ", 1)[0]
            replacement = catalog[catalog.RuthID.eq(replacement_id)].iloc[0]
            rows.append({"Account": item.Account, "Qty": int(item.Qty), "Product": item.Product,
                         "OriginalRuthID": original_id, "ReplacementRuthID": replacement_id,
                         "ReplacementItem": replacement.Item})
    if rows:
        st.download_button("⬇️ REPORT #4 — DOWNLOAD SUGGESTED SUBSTITUTION PICKLIST",
                           pdf_report(rows, rd), f"OOS_Substitutions_{rd}.pdf", "application/pdf",
                           use_container_width=True)
    else:
        st.caption("Choose and approve at least one substitution to enable the PDF.")
