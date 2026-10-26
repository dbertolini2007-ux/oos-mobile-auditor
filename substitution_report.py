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
                "RuthID": data[id_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip(),
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
    return out[out.RuthID.str.fullmatch(r"\d+") & out.Item.ne("")].drop_duplicates("RuthID")


def best_match(description, catalog):
    key = clean(description)
    scores = catalog.copy()
    scores["Score"] = scores.Item.map(lambda x: SequenceMatcher(None, key, clean(x)).ratio())
    scores = scores.sort_values("Score", ascending=False)
    return scores.iloc[0] if len(scores) else None


def candidates(item, catalog, excluded, account_picklist=None):
    orig = best_match(item, catalog)
    if orig is None:
        return None, []
    pool = catalog[~catalog.RuthID.isin(excluded | {orig.RuthID})].copy()
    if pool.empty:
        return orig, []

    def size(text):
        match = re.search(r"(\d+(?:\.\d+)?)\s*(OZ|OUNCE|OUNCES|FL OZ|ML|G|GRAM|GRAMS|LB)\b", clean(text))
        if not match:
            return None
        value, unit = float(match.group(1)), match.group(2)
        if unit in ("OUNCE", "OUNCES", "FL OZ"):
            unit = "OZ"
        if unit in ("GRAM", "GRAMS"):
            unit = "G"
        return value, unit

    original_words = set(clean(orig.Item).split())
    original_size = size(orig.Item)
    def score(row):
        words = set(clean(row.Item).split())
        overlap = len(original_words & words) / max(1, len(original_words | words))
        similarity = SequenceMatcher(None, clean(orig.Item), clean(row.Item)).ratio()
        same_category = bool(clean(orig.Category) and clean(row.Category) == clean(orig.Category))
        candidate_size = size(row.Item)
        size_bonus = 0
        if original_size and candidate_size and original_size[1] == candidate_size[1]:
            ratio = min(original_size[0], candidate_size[0]) / max(original_size[0], candidate_size[0])
            size_bonus = 0.30 * ratio
        elif original_size and candidate_size:
            size_bonus = -0.10
        return (0.50 * overlap + 0.25 * similarity +
                (0.25 if same_category else -0.15) + size_bonus)

    pool["MatchScore"] = pool.apply(score, axis=1)
    same_category = pool[pool.Category.map(clean).eq(clean(orig.Category))] if clean(orig.Category) else pool
    ranked = same_category.sort_values("MatchScore", ascending=False).head(5).copy()
    if ranked.empty:
        return orig, []
    ordered = {clean(x) for x in (account_picklist or []) if clean(x)}
    ranked["AlreadyOnPicklist"] = ranked.Item.map(lambda x: clean(x) in ordered)
    ranked = ranked.sort_values(["AlreadyOnPicklist", "MatchScore"], ascending=[True, False])
    return orig, ranked.to_dict("records")

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


def render_report_4(df, rd, data=None, verified_outs=None):
    st.markdown("### Report #4 — Suggested Substitution Picklist (Optional)")
    st.caption("Uses Verified Outs Column A. Review suggested replacements, location, and quantity before printing.")
    upload = st.file_uploader("Product catalog (ITEMS.xlsx) — upload once per session", type=["xlsx"], key="sub_catalog")
    if upload is None:
        st.download_button("⬇️ REPORT #4 — SUGGESTED SUBSTITUTION PICKLIST", data=b"", file_name="OOS_Substitutions.pdf", mime="application/pdf", disabled=True)
        st.info("Report #4 is not ready: upload ITEMS.xlsx above to generate and approve substitutions.")
        return
    try:
        catalog = catalog_from_excel(upload)
    except Exception as e:
        st.error(f"Cannot read catalog: {e}")
        return
    # Verified Outs column A is the source of truth, independent of the
    # confidence-based audit status. Do not promote FALSE OOS rows.
    verified_names = [str(x).strip() for x in (verified_outs or []) if str(x).strip()]
    audit_rows = df[df["Status"].eq("VERIFIED OOS")].copy() if len(df) else pd.DataFrame()
    verified_rows = []
    for name in verified_names:
        matching = audit_rows[audit_rows.Product.map(
            lambda x: SequenceMatcher(None, clean(x), clean(name)).ratio() >= 0.75
        )] if not audit_rows.empty else pd.DataFrame()
        if not matching.empty:
            for _, line in matching.iterrows():
                verified_rows.append({"Account": line["Account"], "Picker": line.get("Picker", ""),
                                      "Product": name, "Qty": max(1, int(line["Qty"]))})
        else:
            verified_rows.append({"Account": "WAREHOUSE / LOCATION TO CONFIRM",
                                  "Picker": "", "Product": name, "Qty": 1})
    verified = pd.DataFrame(verified_rows)
    if verified.empty:
        st.download_button("⬇️ REPORT #4 — SUGGESTED SUBSTITUTION PICKLIST",
                           data=b"", file_name="OOS_Substitutions.pdf",
                           mime="application/pdf", disabled=True)
        st.info("No products found in Verified Outs Column A. Upload a verified-outs list and run the audit.")
        return
    if any(verified.Account.eq("WAREHOUSE / LOCATION TO CONFIRM")):
        st.info("Items not matched to a verified pick-sheet line are grouped under WAREHOUSE / LOCATION TO CONFIRM with quantity 1. Review quantities and locations before use.")
    excluded = set()
    for product in verified.Product.unique():
        original = best_match(product, catalog)
        if original is not None:
            excluded.add(str(original.RuthID))
    account_products = {}
    for order in (data or {}).get("orders", []):
        account = clean(order.get("account", "Unknown"))
        account_products.setdefault(account, set()).update(
            clean(x) for x in (order.get("all_products") or []) if clean(x)
        )
    rows = []
    for n, item in enumerate(verified.itertuples(index=False)):
        orig, choices = candidates(item.Product, catalog, excluded, account_products.get(clean(item.Account), set()))
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
    st.download_button("⬇️ REPORT #4 — SUGGESTED SUBSTITUTION PICKLIST",
                       pdf_report(rows, rd) if rows else b"", f"OOS_Substitutions_{rd}.pdf", "application/pdf",
                       disabled=not bool(rows), use_container_width=True)
    if not rows:
        st.caption("Choose and approve at least one substitution to enable the PDF.")
