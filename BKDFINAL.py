import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import re
import io
import os
import streamlit.components.v1 as components

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle
)

# =========================
#         UI
# =========================
st.set_page_config(page_title="Extracteur PMC", layout="wide")
st.title("📦 Extracteur de données PMC")

st.markdown("""
Dépose ton fichier manifest (PDF) : le script extrait automatiquement :
- Infos vol + PMC
- Total pièces + détail par AWB (n°, pièces, poids)
- 📄 Export PDF (Résumé PMC), 📊 Stats PMC
""")

uploaded_file = st.file_uploader("📤 Dépose ton fichier PDF ici :", type="pdf")

st.markdown("**Colle ici les PMC à extraire (un par ligne) :**")
pmc_filter_text = st.text_area("PMC à extraire", height=150)
filtered_pmcs = [line.strip() for line in pmc_filter_text.splitlines() if line.strip()]

# =========================
#     Constantes / Regex
# =========================

PMC_PREFIXES = [
    "PMC", "PEB", "PGE", "PGA", "DQF", "PRA", "PAG", "PUB",
    "MPC", "AKH", "QKE", "AKE", "PYB", "AAD",
    "P1P", "PIP", "PAJ", "HMJ"
]

PMC_PATTERN = re.compile(
    rf"^(?:BULK|(?:({'|'.join(map(re.escape, PMC_PREFIXES))})[A-Z0-9]+|PLA\d+[A-Z0-9]*))$"
)

IATA_PAIR = re.compile(r"\b([A-Z]{3})\s*-\s*([A-Z]{3})\b")

HEADER_EXCLUDE = re.compile(
    r"(Flight No\./Date|Point of Loading|Arr\. Date|Import Check Manifest|\*[A-Z0-9\-]+)"
)

# =========================
#   Extraction principale
# =========================
def extract_manifest_with_pcs_awb(pdf_path):
    doc = fitz.open(pdf_path)
    lines_preview = []
    for i in range(min(3, len(doc))):
        lines_preview.extend(doc[i].get_text().splitlines())

    point_of_loading = "UNKNOWN"
    flight_no = "UNKNOWN"
    for i, line in enumerate(lines_preview):
        if "Point of Loading:" in line and i - 1 >= 0:
            point_of_loading = lines_preview[i - 1].strip()
        if "Flight No./Date:" in line and i - 1 >= 0:
            flight_line = lines_preview[i - 1].strip()
            match = re.match(r"^([A-Z0-9]+)", flight_line)
            if match:
                flight_no = match.group(1)

    data = []
    current_pmc = None
    weights = []
    pcs_list = []
    awb_list = []

    for page in doc[:min(50, len(doc))]:
        blocks = page.get_text("blocks")
        lines = []
        for block in blocks:
            text = (block[4] or "").strip()
            if text:
                lines.extend(text.splitlines())

        i = 0
        while i < len(lines):
            line = (lines[i] or "").strip()
            token = line.split()[0] if line else ""

            if token and PMC_PATTERN.match(token):
                if current_pmc:
                    total_weight = sum(weights) if weights else 0
                    total_pcs = sum(pcs_list) if pcs_list else 0
                    data.append([
                        point_of_loading,
                        flight_no,
                        current_pmc,
                        f"{total_weight:.1f}".replace(".", ","),
                        total_pcs,
                        "\n".join(awb_list),
                        "\n".join(str(p) for p in pcs_list),
                        "\n".join(f"{w:.1f}".replace(".", ",") for w in weights),
                        len(awb_list)
                    ])
                current_pmc = token
                weights = []
                pcs_list = []
                awb_list = []
                i += 1
                continue

            if re.match(r'^\d{3}-\d{8}$', line) and i + 2 < len(lines):
                awb = line.strip()
                pcs_line = (lines[i + 1] or "").strip()
                weight_line = (lines[i + 2] or "").strip()
                try:
                    pcs_val = int(pcs_line.split('/')[0])
                    weight_val = float(weight_line.replace(",", ""))
                    awb_list.append(awb)
                    pcs_list.append(pcs_val)
                    weights.append(weight_val)
                    i += 3
                    continue
                except:
                    pass

            match_inline = re.match(r'^(\d{3}-\d{8})\s+(\d+)(?:/\d+)?\s+([\d.,]+)', line)
            if match_inline:
                try:
                    awb = match_inline.group(1)
                    pcs_val = int(match_inline.group(2))
                    weight_val = float(match_inline.group(3).replace(",", ""))
                    awb_list.append(awb)
                    pcs_list.append(pcs_val)
                    weights.append(weight_val)
                    i += 1
                    continue
                except:
                    pass

            i += 1

    if current_pmc:
        total_weight = sum(weights) if weights else 0
        total_pcs = sum(pcs_list) if pcs_list else 0
        data.append([
            point_of_loading,
            flight_no,
            current_pmc,
            f"{total_weight:.1f}".replace(".", ","),
            total_pcs,
            "\n".join(awb_list),
            "\n".join(str(p) for p in pcs_list),
            "\n".join(f"{w:.1f}".replace(".", ",") for w in weights),
            len(awb_list)
        ])

    df = pd.DataFrame(data, columns=[
        "Point of Loading", "Flight No", "PMC No", "Poids brut (kg)",
        "Total Pièces", "Liste des AWB", "Pièces par AWB", "Poids par AWB", "Nombre AWB"
    ])
    df = df.sort_values("Total Pièces").reset_index(drop=True)
    return df

# =========================
#     Helpers DEST (DST uniquement)
# =========================
def _collect_pdf_lines(pdf_path):
    doc = fitz.open(pdf_path)
    pages_lines = []
    for page in doc:
        blocks = page.get_text("blocks")
        lines = []
        for b in blocks:
            txt = (b[4] or "").strip()
            if txt:
                lines.extend(txt.splitlines())
        pages_lines.append(lines)
    return pages_lines

def build_awb_destination_map(pdf_path):
    """
    Associe chaque AWB au code destination (DST, 3 lettres) :
    1) Cherche ORG-DST sur la même ligne (priorité à la partie après l'AWB),
    2) Sinon scanne 3 lignes suivantes (en ignorant les en-têtes),
    3) Sinon 2 lignes précédentes (en ignorant les en-têtes).
    Renvoie uniquement le DST (2e code).
    """
    pages_lines = _collect_pdf_lines(pdf_path)
    awb_dest = {}

    for lines in pages_lines:
        for i, line in enumerate(lines):
            for m_awb in re.finditer(r"\b(\d{3}-\d{8})\b", line):
                awb = m_awb.group(1)
                if awb in awb_dest:
                    continue

                dst_code = None
                post = line[m_awb.end():]
                m_dst = IATA_PAIR.search(post)

                if not m_dst:
                    m_dst = IATA_PAIR.search(line)

                if not m_dst:
                    for j in range(1, 4):
                        if i + j >= len(lines):
                            break
                        ln = lines[i + j]
                        if HEADER_EXCLUDE.search(ln):
                            continue
                        m_dst = IATA_PAIR.search(ln)
                        if m_dst:
                            break

                if not m_dst:
                    for j in range(1, 3):
                        if i - j < 0:
                            break
                        ln = lines[i - j]
                        if HEADER_EXCLUDE.search(ln):
                            continue
                        m_dst = IATA_PAIR.search(ln)
                        if m_dst:
                            break

                if m_dst:
                    dst_code = m_dst.group(2)

                awb_dest[awb] = dst_code or ""

    return awb_dest

# =========================
#   PDF Résumé PMC uniquement
# =========================
def generate_summary_pdf(dataframe, source_pdf_path):
    """
    Construit un PDF contenant UNIQUEMENT le résumé PMC :
    - tableau zébré
    - colonne Localisation vide
    - colonne Destination
    - colonne Total AWB
    """
    awb_totals = {}
    for _, row in dataframe.iterrows():
        awb_list = str(row.get("Liste des AWB", "") or "").split("\n")
        pcs_list = str(row.get("Pièces par AWB", "") or "").split("\n")
        for awb, pcs in zip(awb_list, pcs_list):
            awb = awb.strip()
            if not awb:
                continue
            try:
                pcs_val = int(str(pcs).strip().split("/")[0])
            except:
                pcs_val = 0
            awb_totals[awb] = awb_totals.get(awb, 0) + pcs_val

    awb_dest_map = build_awb_destination_map(source_pdf_path)

    df_pdf = dataframe.copy()

    if "Liste des AWB" in df_pdf.columns:
        insert_idx = list(df_pdf.columns).index("Liste des AWB") + 1
    else:
        insert_idx = len(df_pdf.columns)

    df_pdf.insert(insert_idx, "Localisation", "")
    df_pdf.insert(insert_idx + 1, "Destination", "")

    def _row_total_awb_strings(row):
        awbs = str(row.get("Liste des AWB", "") or "").split("\n")
        return "\n".join(str(awb_totals.get(a.strip(), 0)) for a in awbs if a.strip())

    def _row_destinations(row):
        awbs = str(row.get("Liste des AWB", "") or "").split("\n")
        return "\n".join(awb_dest_map.get(a.strip(), "") for a in awbs if a.strip())

    if "Pièces par AWB" in df_pdf.columns:
        insert_idx2 = list(df_pdf.columns).index("Pièces par AWB") + 1
    else:
        insert_idx2 = len(df_pdf.columns)

    df_pdf.insert(insert_idx2, "Total AWB", df_pdf.apply(_row_total_awb_strings, axis=1))

    if "Destination" in df_pdf.columns:
        df_pdf["Destination"] = df_pdf.apply(_row_destinations, axis=1)

    drop_cols = [c for c in ["Point of Loading", "Flight No", "Nombre AWB", "Poids brut (kg)"] if c in df_pdf.columns]
    df_pdf = df_pdf.drop(columns=drop_cols)

    cols = list(df_pdf.columns)
    desired_prefix = [
        c for c in [
            "PMC No", "Liste des AWB", "Localisation", "Destination",
            "Pièces par AWB", "Total AWB", "Poids par AWB"
        ] if c in cols
    ]
    others = [c for c in cols if c not in desired_prefix + ["Total Pièces"]]
    tail = ["Total Pièces"] if "Total Pièces" in cols else []
    new_order = desired_prefix + others + tail
    df_pdf = df_pdf[new_order]

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    story = []

    data_table = [list(df_pdf.columns)] + df_pdf.astype(str).values.tolist()
    table = Table(data_table, repeatRows=1)
    table_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BOX", (0, 0), (-1, -1), 0.25, colors.black),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ])
    table.setStyle(table_style)

    story.append(table)

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# =========================
#   Helper texte résumé PMC
# =========================
def build_pmc_bulk_summary_line(dataframe):
    pmc_values = dataframe["PMC No"].astype(str).str.strip().tolist()

    has_bulk = any(val.upper() == "BULK" for val in pmc_values)
    pmc_count = sum(1 for val in pmc_values if val.upper() != "BULK")
    total_pcs = dataframe["Total Pièces"].astype(int).sum()

    if has_bulk:
        return f"{pmc_count} PMC + BULK ({total_pcs} PCS)"
    return f"{pmc_count} PMC ({total_pcs} PCS)"

# =========================
#   Bouton copier + texte
# =========================
def render_copy_button_left(text_to_copy):
    safe_text = (
        text_to_copy
        .replace("\\", "\\\\")
        .replace("`", "\\`")
        .replace("$", "\\$")
    )

    components.html(
        f"""
        <div style="display:flex; align-items:center; gap:10px; margin-top:2px; margin-bottom:6px;">
            <button
                onclick="navigator.clipboard.writeText(`{safe_text}`); this.innerText='Copié';"
                style="
                    display:inline-flex;
                    align-items:center;
                    justify-content:center;
                    background-color:rgb(255, 75, 75);
                    color:white;
                    border:none;
                    border-radius:8px;
                    padding:0.45rem 0.9rem;
                    font-size:14px;
                    font-weight:600;
                    cursor:pointer;
                    line-height:1.2;
                "
            >
                Copier
            </button>

            <span style="
                font-size:16px;
                font-weight:600;
                color:white;
                background-color:#1e1e1e;
                padding:6px 10px;
                border-radius:6px;
                font-family:sans-serif;
            ">
                {text_to_copy}
            </span>
        </div>
        """,
        height=55,
    )

# =========================
#        App logic
# =========================
if uploaded_file:
    file_name = os.path.splitext(uploaded_file.name)[0]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.read())
        source_pdf_path = tmp_file.name
        df_result = extract_manifest_with_pcs_awb(source_pdf_path)

    if filtered_pmcs:
        df_result = df_result[df_result["PMC No"].isin(filtered_pmcs)]

    st.success("✅ Extraction terminée avec succès !")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📋 Résultats détaillés par PMC")
        st.dataframe(df_result)

        total_pcs = df_result["Total Pièces"].astype(int).sum()
        total_kg = df_result["Poids brut (kg)"].str.replace(",", ".").astype(float).sum()

        st.markdown(f"**🧮 Total pièces : {total_pcs}**")
        st.markdown(f"**⚖️ Poids total (kg) : {total_kg:,.1f}**".replace(",", " ").replace(".", ","))

        summary_line = build_pmc_bulk_summary_line(df_result)
        render_copy_button_left(summary_line)

        summary_pdf_bytes = generate_summary_pdf(df_result, source_pdf_path)
        st.download_button(
            "📄 Télécharger le PDF résumé",
            data=summary_pdf_bytes,
            file_name=f"{file_name}+RESUME.pdf",
            mime="application/pdf"
        )

    with col2:
        st.subheader("📊 Statistiques sur les pièces par PMC")

        bin_ranges = {
            "< 50": (0, 49),
            "50 - 99": (50, 99),
            "100 - 149": (100, 149),
            "150 - 199": (150, 199),
            "200 - 249": (200, 249),
            "≥ 250": (250, float("inf"))
        }

        stats = {label: 0 for label in bin_ranges}
        for pcs in df_result["Total Pièces"].astype(int):
            for label, (low, high) in bin_ranges.items():
                if low <= pcs <= high:
                    stats[label] += 1
                    break

        df_stats = pd.DataFrame(list(stats.items()), columns=["Tranche de pièces", "Nombre de PMC"])
        total_pmc = int(df_result.shape[0])
        df_stats.loc[len(df_stats)] = ["Total", total_pmc]

        st.table(df_stats)