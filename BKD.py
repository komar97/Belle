import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import re
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

st.set_page_config(page_title="Extracteur PMC", layout="wide")
st.title("📦 Extracteur de données PMC")

st.markdown("""
Dépose ton fichier manifest (PDF) : le script extrait automatiquement :
- Infos vol + PMC
- Total pièces + détail par AWB (n°, pièces, poids)
- 📄 Export PDF, 📊 Stats PMC

✍️ Colle ici une liste de PMC à extraire si tu veux en filtrer certains.
""")

uploaded_file = st.file_uploader("📤 Dépose ton fichier PDF ici :", type="pdf")

st.markdown("**Colle ici les PMC à extraire (un par ligne) :**")
pmc_filter_text = st.text_area("PMC à extraire", height=150)
filtered_pmcs = [line.strip() for line in pmc_filter_text.splitlines() if line.strip()]

# Expression stricte : préfixes + suite alphanumérique OU "BULK"
PMC_PATTERN = re.compile(
    r"^(" +
    r"BULK" +
    r"|" +
    r"(?:PMC|PEB|PGE|PGA|PRA|PAG|PUB|PLA|MPC|QKE|AKE|PYB|PIP|PAJ|HMJ)[A-Z0-9]+" +
    r")$"
)

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
            text = block[4].strip()
            if text:
                lines.extend(text.splitlines())

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            if PMC_PATTERN.match(line):
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
                current_pmc = line.split()[0]
                weights = []
                pcs_list = []
                awb_list = []
                i += 1
                continue

            if re.match(r'^\d{3}-\d{8}$', line) and i + 2 < len(lines):
                awb = line.strip()
                pcs_line = lines[i + 1].strip()
                weight_line = lines[i + 2].strip()

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

def generate_pdf(dataframe):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)

    # Supprimer colonnes spécifiques pour le PDF
    df_pdf = dataframe.drop(columns=["Point of Loading", "Flight No", "Nombre AWB"])

    data = [list(df_pdf.columns)] + df_pdf.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BOX", (0, 0), (-1, -1), 0.25, colors.black),
    ])
    for row in range(1, len(data)):
        bg_color = colors.HexColor("#C0C0C0") if row % 2 == 0 else colors.white
        style.add("BACKGROUND", (0, row), (-1, row), bg_color)
    table.setStyle(style)
    doc.build([table])
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

if uploaded_file:
    file_name = os.path.splitext(uploaded_file.name)[0]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.read())
        df_result = extract_manifest_with_pcs_awb(tmp_file.name)

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

        pdf_bytes = generate_pdf(df_result)
        st.download_button(
            label="📄 Télécharger en PDF",
            data=pdf_bytes,
            file_name=f"{file_name}_RESUME.pdf",
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
        st.table(df_stats)
