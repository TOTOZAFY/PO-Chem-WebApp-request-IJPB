import os
import re
import ssl
import smtplib
from datetime import date
from email.message import EmailMessage
from io import BytesIO

import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(
    page_title="Projet / Project request | OV-Chimie-Métabolisme",
    page_icon="🧪",
    layout="wide",
)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_E-GES-01.docx")
STEPS = [
    "Contacts",
    "Description du projet",
    "Échantillons",
    "Résultats & bibliographie",
    "Informations complémentaires",
    "Vérification & envoi",
]

st.markdown(
    """
    <style>
    .block-container {max-width: 1050px; padding-top: 2rem; padding-bottom: 3rem;}
    .hero {padding: 1.3rem 1.5rem; border-radius: 14px;
           background: linear-gradient(120deg,#edf5fb,#f4f7f2); margin-bottom: 1rem;}
    .hero h1 {margin: 0; color: #183b56; font-size: 1.8rem;}
    .hero p {margin: .4rem 0 0; color: #526575;}
    </style>
    <div class="hero">
      <h1>OV-Chimie-Métabolisme</h1>
      <p>Formulaire de demande d'analyse / Analysis project request form — E-GES-01</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write(
    "Complétez les rubriques ci-dessous. À la dernière étape, vous pourrez "
    "télécharger le document Word. L'envoi par e-mail est facultatif et nécessite "
    "une configuration SMTP par l'administrateur."
)

if not os.path.isfile(TEMPLATE_PATH):
    st.error(
        "Le modèle Word template_E-GES-01.docx est introuvable. "
        "Vérifiez qu'il se trouve dans le même dossier que app.py."
    )
    st.stop()

if "step" not in st.session_state:
    st.session_state.step = 0
if "data" not in st.session_state:
    st.session_state.data = {}

data = st.session_state.data


def text_field(key, label, **kwargs):
    data[key] = st.text_input(
        label, value=data.get(key, ""), key=f"field_{key}", **kwargs
    )


def text_area(key, label, height=150):
    data[key] = st.text_area(
        label, value=data.get(key, ""), height=height, key=f"field_{key}"
    )


def add_value(cell, value, *, font_size=9, bold=False):
    """Append a value without deleting the original template label."""
    if value is None or str(value).strip() == "":
        return
    paragraph = cell.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(2)
    run = paragraph.add_run(str(value).strip())
    run.font.size = Pt(font_size)
    run.bold = bold


def add_to_existing_paragraph(paragraph, value, *, font_size=9, bold=False):
    if value is None or str(value).strip() == "":
        return
    run = paragraph.add_run(" " + str(value).strip())
    run.font.size = Pt(font_size)
    run.bold = bold


def build_document(form_data):
    doc = Document(TEMPLATE_PATH)

    if len(doc.tables) < 2:
        raise ValueError("Le modèle Word ne contient pas le tableau de formulaire attendu.")

    # Table 1: page 1 contact fields and project sections.
    table = doc.tables[1]

    # Contact details. The template uses merged cells in several rows.
    add_to_existing_paragraph(table.cell(0, 0).paragraphs[0], form_data.get("project_manager"))
    add_to_existing_paragraph(table.cell(0, 2).paragraphs[0], form_data.get("group"))

    add_to_existing_paragraph(table.cell(1, 0).paragraphs[0], form_data.get("date"))
    add_to_existing_paragraph(table.cell(1, 1).paragraphs[0], form_data.get("manager_phone"))
    add_to_existing_paragraph(table.cell(1, 3).paragraphs[0], form_data.get("manager_email"))

    add_to_existing_paragraph(table.cell(2, 0).paragraphs[0], form_data.get("admin_staff"))
    add_to_existing_paragraph(table.cell(2, 1).paragraphs[0], form_data.get("admin_phone"))
    add_to_existing_paragraph(table.cell(2, 3).paragraphs[0], form_data.get("admin_email"))

    # Projet / Project description: insert into the blank paragraphs above the platform notice.
    description_cell = table.cell(3, 0)
    description_paragraphs = description_cell.paragraphs
    description = form_data.get("project_description", "").strip()
    if description:
        blank_paragraphs = [
            p for p in description_paragraphs[2:-1] if not p.text.strip()
        ]
        target = blank_paragraphs[0] if blank_paragraphs else description_cell.add_paragraph()
        target.paragraph_format.space_before = Pt(4)
        target.paragraph_format.space_after = Pt(4)
        run = target.add_run(description)
        run.font.size = Pt(9)

    # Sample details: use the existing blank paragraph under the label.
    samples_cell = table.cell(4, 0)
    samples = form_data.get("samples", "").strip()
    if samples:
        target = next((p for p in samples_cell.paragraphs if not p.text.strip()), None)
        if target is None:
            target = samples_cell.add_paragraph()
        run = target.add_run(samples)
        run.font.size = Pt(9)

    # Preliminary results and bibliography.
    results_cell = table.cell(5, 0)
    add_value(results_cell, form_data.get("preliminary", ""))
    bibliography = form_data.get("bibliography", "").strip()
    if bibliography:
        add_value(results_cell, "Bibliographie : " + bibliography)

    # Projet / Project name and YES/NO choices. Original labels and page layout are retained.
    project_cell = table.cell(6, 0)
    add_to_existing_paragraph(project_cell.paragraphs[0], form_data.get("project_name"), bold=True)
    funding = form_data.get("funding", "NO")
    add_to_existing_paragraph(project_cell.paragraphs[1], f"Réponse : {funding}", bold=True)

    guidance_cell = table.cell(7, 0)
    add_to_existing_paragraph(guidance_cell.paragraphs[0], f"Réponse : {form_data.get('guidance', 'NO')}", bold=True)

    # Update the reference in the small header table, leaving page 2 untouched.
    ref_name = re.sub(
        r"[^A-Za-z0-9]+", "_",
        form_data.get("project_manager", "NOM")
    ).strip("_").upper() or "NOM"
    reference = f"E-GES-01_{date.today().strftime('%d%m%y')}_{ref_name}"
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if "E-GES-01_AAMMJJ_NOM" in paragraph.text:
                        for run in paragraph.runs:
                            run.text = run.text.replace("E-GES-01_AAMMJJ_NOM", reference)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue(), reference + ".docx"


# Form navigation and content.
step = st.session_state.step
st.progress((step + 1) / len(STEPS), text=f"Étape {step + 1}/{len(STEPS)} — {STEPS[step]}")

with st.container(border=True):
    if step == 0:
        st.subheader("1. Informations générales")
        left, right = st.columns(2)
        with left:
            text_field("project_manager", "Responsable du projet *")
            text_field("group", "Groupe / laboratoire")
            selected_date = st.date_input(
                "Date / Date",
                value=date.fromisoformat(data["date"]) if data.get("date") else date.today(),
            )
            data["date"] = selected_date.isoformat()
            text_field("manager_phone", "Téléphone")
            text_field("manager_email", "E-mail du responsable")
        with right:
            text_field("admin_staff", "Personnel administratif")
            text_field("admin_phone", "Téléphone du personnel administratif")
            text_field("admin_email", "E-mail du personnel administratif")
        st.caption("* Champ recommandé.")

    elif step == 1:
        st.subheader("2. Description du projet")
        st.caption("Questions scientifiques et techniques à traiter, molécules à analyser.")
        text_area("project_description", "Description scientifique et technique", 230)

    elif step == 2:
        st.subheader("3. Nature et nombre des échantillons")
        st.caption(
            "Indiquez notamment les références du procédé analytique, les solvants "
            "et les formules chimiques si disponibles."
        )
        text_area("samples", "Nature et nombre des échantillons / informations analytiques", 220)

    elif step == 3:
        st.subheader("4. Résultats préliminaires et bibliographie")
        text_area("preliminary", "Résultats déjà obtenus et éléments utiles à l'évaluation du projet", 160)
        text_area("bibliography", "Résumé bibliographique et références des articles pertinents", 160)

    elif step == 4:
        st.subheader("5. Informations complémentaires")
        text_field("project_name", "Nom du projet *")
        funding_options = ["YES", "NO"]
        guidance_options = ["YES", "NO"]
        data["funding"] = st.radio(
            "Financement (uniquement pour les chercheurs IJPB)",
            funding_options,
            horizontal=True,
            index=funding_options.index(data.get("funding", "NO")),
        )
        data["guidance"] = st.radio(
            "Souhaitez-vous un accompagnement scientifique ?",
            guidance_options,
            horizontal=True,
            index=guidance_options.index(data.get("guidance", "NO")),
        )
        st.info(
            "La plateforme peut vous mettre en relation avec des scientifiques de l'IJPB "
            "pour aider à interpréter les résultats. Une co-signature / co-autorat pourra "
            "alors être requise."
        )

    else:
        st.subheader("6. Vérification, génération et envoi")
        summary = [
            ("Responsable", data.get("project_manager", "")),
            ("Groupe", data.get("group", "")),
            ("Date / Date", data.get("date", "")),
            ("E-mail", data.get("manager_email", "")),
            ("Projet", data.get("project_name", "")),
            ("Financement IJPB", data.get("funding", "NO")),
            ("Accompagnement scientifique", data.get("guidance", "NO")),
        ]
        for label, value in summary:
            st.markdown(f"**{label} :** {value or '—'}")

        with st.expander("Afficher toutes les réponses"):
            for key, value in data.items():
                st.markdown(f"**{key.replace('_', ' ').capitalize()} :** {value}")

        st.divider()
        recipient = st.text_input("Envoyer le document à cette adresse e-mail", value="ov-chimie@inrae.fr")
        subject = st.text_input(
            "Objet du message",
            value=f"Demande de projet – {data.get('project_name', 'E-GES-01')}",
        )
        sender_name = st.text_input(
            "Nom de l'expéditeur (facultatif)",
            value=data.get("project_manager", ""),
        )

        if not data.get("project_manager") or not data.get("project_name"):
            st.warning("Renseignez au minimum le responsable et le nom du projet avant la génération.")
        else:
            try:
                document_bytes, filename = build_document(data)
                st.download_button(
                    "📄 Télécharger le document Word",
                    data=document_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"Impossible de générer le document Word : {exc}")
                document_bytes = None
                filename = None

            smtp_configured = "smtp" in st.secrets
            if not smtp_configured:
                st.caption(
                    "L'envoi automatique par e-mail n'est pas encore configuré. "
                    "Le téléchargement Word reste disponible."
                )

            if st.button(
                "✉️ Envoyer le document par e-mail",
                use_container_width=True,
                disabled=(not smtp_configured or document_bytes is None),
            ):
                try:
                    cfg = st.secrets["smtp"]
                    message = EmailMessage()
                    message["Subject"] = subject
                    message["From"] = cfg["sender"]
                    message["To"] = recipient
                    message.set_content(
                        f"Bonjour,\n\nVeuillez trouver ci-joint la demande de projet E-GES-01."
                        f"\n\nCordialement,\n{sender_name}"
                    )
                    message.add_attachment(
                        document_bytes,
                        maintype="application",
                        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                        filename=filename,
                    )
                    with smtplib.SMTP(cfg["host"], int(cfg.get("port", 587)), timeout=20) as server:
                        server.starttls(context=ssl.create_default_context())
                        server.login(cfg["username"], cfg["password"])
                        server.send_message(message)
                    st.success(f"Document envoyé à {recipient}.")
                except Exception as exc:
                    st.error(f"Envoi impossible. Le document reste téléchargeable. Détail : {exc}")

# Navigation buttons.
st.divider()
if step < len(STEPS) - 1:
    previous_col, next_col = st.columns(2)
    with previous_col:
        if st.button("← Précédent", disabled=(step == 0), use_container_width=True):
            st.session_state.step -= 1
            st.rerun()
    with next_col:
        if st.button("Suivant →", use_container_width=True):
            st.session_state.step += 1
            st.rerun()
else:
    if st.button("← Modifier les réponses", use_container_width=True):
        st.session_state.step -= 1
        st.rerun()
