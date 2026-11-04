import os
import re
import ssl
import smtplib
import json
import uuid
from datetime import date, datetime
from email.message import EmailMessage
from io import BytesIO
from urllib.parse import quote

import requests
import streamlit as st
from docx import Document
from docx.shared import Pt

st.set_page_config(
    page_title="PO-Chem-WebApp-Request | OV-Chimie-Métabolisme",
    page_icon="🧪",
    layout="wide",
)

APP_NAME = "PO-Chem-WebApp-Request"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_E-GES-01.docx")
ADMIN_EMAIL = "ov-chimie@inrae.fr"
CLOUD_FOLDER = "MyDrive/E-GES-01/Demandes_Word"
METADATA_PATH = "MyDrive/E-GES-01/demandes_metadata.json"

TRANSLATIONS = {
    "fr": {
        "language": "Langue / Language", "form": "Formulaire de demande",
        "admin": "Espace administrateur", "intro": "Complétez les rubriques ci-dessous. "
        "À la dernière étape, votre demande sera enregistrée dans Nextcloud et l'administrateur sera notifié.",
        "steps": ["Contacts", "Description du projet", "Échantillons", "Résultats & bibliographie",
                  "Informations complémentaires", "Vérification & soumission"],
        "general": "1. Informations générales", "manager": "Responsable du projet *",
        "group": "Groupe / laboratoire", "date": "Date", "phone": "Téléphone",
        "manager_email": "E-mail du responsable", "admin_staff": "Personnel administratif",
        "admin_phone": "Téléphone du personnel administratif", "admin_email": "E-mail du personnel administratif",
        "recommended": "* Champ obligatoire.", "description": "2. Description du projet",
        "description_help": "Questions scientifiques et techniques à traiter, molécules à analyser.",
        "description_field": "Description scientifique et technique",
        "samples_title": "3. Nature et nombre des échantillons",
        "samples_help": "Indiquez notamment les références du procédé analytique, les solvants et les formules chimiques si disponibles.",
        "samples_field": "Nature et nombre des échantillons / informations analytiques",
        "results_title": "4. Résultats préliminaires et bibliographie",
        "preliminary": "Résultats déjà obtenus et éléments utiles à l'évaluation du projet",
        "bibliography": "Résumé bibliographique et références des articles pertinents",
        "additional": "5. Informations complémentaires", "project_name": "Nom du projet *",
        "funding": "Financement (uniquement pour les chercheurs IJPB)",
        "guidance": "Souhaitez-vous un accompagnement scientifique ?",
        "yes": "Oui", "no": "Non",
        "guidance_info": "La plateforme peut vous mettre en relation avec des scientifiques de l'IJPB pour aider à interpréter les résultats. Une co-signature / co-autorat pourra alors être requise.",
        "review": "6. Vérification et soumission", "all_answers": "Afficher toutes les réponses",
        "previous": "← Précédent", "next": "Suivant →", "edit": "← Modifier les réponses",
        "required": "Renseignez au minimum le responsable, son e-mail et le nom du projet.",
        "submit": "Valider et soumettre la demande", "download": "Télécharger le document Word",
        "submitted": "Demande enregistrée dans Nextcloud. L'administrateur a été notifié.",
        "cloud_error": "Le dépôt dans Nextcloud a échoué. La demande n'a pas été enregistrée :",
        "mail_warning": "Le document est enregistré, mais la notification à l'administrateur n'a pas pu être envoyée :",
        "cloud_not_configured": "La connexion Nextcloud n'est pas configurée dans les secrets Streamlit.",
        "admin_login": "Connexion administrateur", "password": "Mot de passe administrateur",
        "login": "Se connecter", "logout": "Se déconnecter", "wrong_password": "Mot de passe incorrect.",
        "admin_title": "Tableau de bord administrateur", "total": "Demandes reçues",
        "pending": "En attente", "accepted": "Acceptées", "refused": "Refusées",
        "by_status": "Répartition par statut", "by_month": "Demandes par mois",
        "request_list": "Gestion des demandes", "no_requests": "Aucune demande enregistrée.",
        "select_request": "Sélectionner une demande", "status": "Décision / statut",
        "comment": "Commentaire facultatif", "save_decision": "Enregistrer la décision et notifier le demandeur",
        "decision_saved": "Décision enregistrée et notification envoyée au demandeur.",
        "decision_saved_no_mail": "Décision enregistrée, mais l'e-mail au demandeur n'a pas pu être envoyé :",
        "select_email_missing": "Aucune adresse e-mail du demandeur n'est renseignée.",
        "cloud_load_error": "Impossible de lire les données depuis Nextcloud :",
        "not_configured_admin": "Espace administrateur indisponible : configurez ADMIN_PASSWORD dans les secrets Streamlit.",
        "doc_path": "Document dans Nextcloud", "created": "Soumise le", "status_pending": "En attente",
        "status_accepted": "Acceptée", "status_refused": "Refusée",
        "mail_admin_subject": "Nouvelle demande E-GES-01 reçue",
        "mail_user_subject": "Mise à jour de votre demande E-GES-01",
        "mail_admin_body": "Une nouvelle demande E-GES-01 a été soumise.\nProjet : {project}\nResponsable : {manager}\nGroupe : {group}\nDocument stocké dans Nextcloud : {path}",
        "mail_user_body": "Bonjour {manager},\n\nLa demande concernant le projet « {project} » a été mise à jour.\nDécision : {status}\nCommentaire : {comment}\n\nCordialement,\nOV-Chimie-Métabolisme",
        "back_form": "Retour au formulaire",
    },
    "en": {
        "language": "Language / Langue", "form": "Request form",
        "admin": "Administrator area", "intro": "Complete the sections below. "
        "At the final step, your request will be saved to Nextcloud and the administrator will be notified.",
        "steps": ["Contacts", "Project description", "Samples", "Preliminary results & bibliography",
                  "Additional information", "Review & submission"],
        "general": "1. General information", "manager": "Project manager *",
        "group": "Group / laboratory", "date": "Date", "phone": "Telephone",
        "manager_email": "Project manager email", "admin_staff": "Administrative staff",
        "admin_phone": "Administrative staff telephone", "admin_email": "Administrative staff email",
        "recommended": "* Required field.", "description": "2. Project description",
        "description_help": "Scientific and technical questions to be addressed, molecules to be analysed.",
        "description_field": "Scientific and technical description",
        "samples_title": "3. Nature and number of samples",
        "samples_help": "Please include analytical process references, solvents and chemical formulas if available.",
        "samples_field": "Nature and number of samples / analytical information",
        "results_title": "4. Preliminary results and bibliography",
        "preliminary": "Results already obtained and information useful for project assessment",
        "bibliography": "Bibliography summary and references to relevant papers",
        "additional": "5. Additional information", "project_name": "Project name *",
        "funding": "Funding (IJPB researchers only)",
        "guidance": "Would you like scientific guidance?",
        "yes": "Yes", "no": "No",
        "guidance_info": "The platform can connect you with IJPB scientists to help interpret the results. Co-authorship may then be required.",
        "review": "6. Review and submission", "all_answers": "Show all answers",
        "previous": "← Previous", "next": "Next →", "edit": "← Edit answers",
        "required": "Please provide at least the project manager, their email address and the project name.",
        "submit": "Submit request", "download": "Download Word document",
        "submitted": "Request saved to Nextcloud. The administrator has been notified.",
        "cloud_error": "Uploading to Nextcloud failed. The request was not saved:",
        "mail_warning": "The document was saved, but the administrator notification could not be sent:",
        "cloud_not_configured": "Nextcloud connection is not configured in Streamlit secrets.",
        "admin_login": "Administrator login", "password": "Administrator password",
        "login": "Log in", "logout": "Log out", "wrong_password": "Incorrect password.",
        "admin_title": "Administrator dashboard", "total": "Requests received",
        "pending": "Pending", "accepted": "Accepted", "refused": "Rejected",
        "by_status": "Requests by status", "by_month": "Requests per month",
        "request_list": "Request management", "no_requests": "No requests have been recorded.",
        "select_request": "Select a request", "status": "Decision / status",
        "comment": "Optional comment", "save_decision": "Save decision and notify requester",
        "decision_saved": "Decision saved and requester notified.",
        "decision_saved_no_mail": "Decision saved, but the requester email could not be sent:",
        "select_email_missing": "No requester email address is available.",
        "cloud_load_error": "Could not read data from Nextcloud:",
        "not_configured_admin": "Administrator area unavailable: set ADMIN_PASSWORD in Streamlit secrets.",
        "doc_path": "Document in Nextcloud", "created": "Submitted on", "status_pending": "Pending",
        "status_accepted": "Accepted", "status_refused": "Rejected",
        "mail_admin_subject": "New E-GES-01 request received",
        "mail_user_subject": "Update on your E-GES-01 request",
        "mail_admin_body": "A new E-GES-01 request has been submitted.\nProject: {project}\nManager: {manager}\nGroup: {group}\nDocument stored in Nextcloud: {path}",
        "mail_user_body": "Hello {manager},\n\nThe request for project “{project}” has been updated.\nDecision: {status}\nComment: {comment}\n\nBest regards,\nOV-Chimie-Métabolisme",
        "back_form": "Back to request form",
    },
}

def t(key):
    return TRANSLATIONS[st.session_state.get("language", "fr")][key]

def get_secret(name, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return default

def nextcloud_settings():
    nested = get_secret("nextcloud", {})
    url = get_secret("NEXTCLOUD_URL", nested.get("url", "") if hasattr(nested, "get") else "")
    username = get_secret("NEXTCLOUD_USERNAME", nested.get("username", "") if hasattr(nested, "get") else "")
    password = get_secret("NEXTCLOUD_APP_PASSWORD", nested.get("app_password", "") if hasattr(nested, "get") else "")
    if not (url and username and password):
        return None
    url = str(url).rstrip("/")
    if "/remote.php/dav/files/" not in url:
        url = f"{url}/remote.php/dav/files/{quote(str(username), safe='')}"
    return url.rstrip("/") + "/", str(username), str(password)

def webdav_request(method, relative_path, *, data=None, timeout=30):
    config = nextcloud_settings()
    if not config:
        raise RuntimeError(t("cloud_not_configured"))
    base_url, username, password = config
    encoded_path = quote(relative_path.strip("/"), safe="/")
    url = base_url + encoded_path
    response = requests.request(method, url, auth=(username, password), data=data, timeout=timeout)
    return response

def ensure_remote_folder(path):
    config = nextcloud_settings()
    if not config:
        raise RuntimeError(t("cloud_not_configured"))
    parts = path.strip("/").split("/")
    current = ""
    for part in parts:
        current = f"{current}/{part}" if current else part
        response = webdav_request("MKCOL", current)
        if response.status_code not in (201, 405):
            response.raise_for_status()

def load_metadata():
    response = webdav_request("GET", METADATA_PATH)
    if response.status_code == 404:
        return []
    response.raise_for_status()
    try:
        payload = response.json()
        return payload if isinstance(payload, list) else []
    except ValueError:
        raise RuntimeError("Le fichier de métadonnées Nextcloud n'est pas un JSON valide.")

def save_metadata(records):
    response = webdav_request(
        "PUT", METADATA_PATH,
        data=json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8"),
        timeout=30,
    )
    response.raise_for_status()

def send_email(to_address, subject, body, cc=None, reply_to=None):
    cfg = get_secret("smtp", {})
    if not cfg or not all(cfg.get(k) for k in ("host", "username", "password", "sender")):
        raise RuntimeError("SMTP n'est pas configuré dans les secrets Streamlit.")
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = cfg["sender"]
    message["To"] = to_address
    if cc:
        message["Cc"] = cc
    if reply_to:
        message["Reply-To"] = reply_to
    message.set_content(body)
    with smtplib.SMTP(cfg["host"], int(cfg.get("port", 587)), timeout=25) as server:
        server.starttls(context=ssl.create_default_context())
        server.login(cfg["username"], cfg["password"])
        server.send_message(message)

def add_value(cell, value, *, font_size=9, bold=False):
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
    table = doc.tables[1]
    add_to_existing_paragraph(table.cell(0, 0).paragraphs[0], form_data.get("project_manager"))
    add_to_existing_paragraph(table.cell(0, 2).paragraphs[0], form_data.get("group"))
    add_to_existing_paragraph(table.cell(1, 0).paragraphs[0], form_data.get("date"))
    add_to_existing_paragraph(table.cell(1, 1).paragraphs[0], form_data.get("manager_phone"))
    add_to_existing_paragraph(table.cell(1, 3).paragraphs[0], form_data.get("manager_email"))
    add_to_existing_paragraph(table.cell(2, 0).paragraphs[0], form_data.get("admin_staff"))
    add_to_existing_paragraph(table.cell(2, 1).paragraphs[0], form_data.get("admin_phone"))
    add_to_existing_paragraph(table.cell(2, 3).paragraphs[0], form_data.get("admin_email"))

    description_cell = table.cell(3, 0)
    description_paragraphs = description_cell.paragraphs
    description = form_data.get("project_description", "").strip()
    if description:
        blank_paragraphs = [p for p in description_paragraphs[2:-1] if not p.text.strip()]
        target = blank_paragraphs[0] if blank_paragraphs else description_cell.add_paragraph()
        target.paragraph_format.space_before = Pt(4)
        target.paragraph_format.space_after = Pt(4)
        target.add_run(description).font.size = Pt(9)

    samples_cell = table.cell(4, 0)
    samples = form_data.get("samples", "").strip()
    if samples:
        target = next((p for p in samples_cell.paragraphs if not p.text.strip()), None)
        if target is None:
            target = samples_cell.add_paragraph()
        target.add_run(samples).font.size = Pt(9)

    results_cell = table.cell(5, 0)
    add_value(results_cell, form_data.get("preliminary", ""))
    bibliography = form_data.get("bibliography", "").strip()
    if bibliography:
        add_value(results_cell, "Bibliographie / Bibliography: " + bibliography)

    project_cell = table.cell(6, 0)
    add_to_existing_paragraph(project_cell.paragraphs[0], form_data.get("project_name"), bold=True)
    add_to_existing_paragraph(project_cell.paragraphs[1], f"Réponse / Answer: {form_data.get('funding', 'NO')}", bold=True)
    guidance_cell = table.cell(7, 0)
    add_to_existing_paragraph(guidance_cell.paragraphs[0], f"Réponse / Answer: {form_data.get('guidance', 'NO')}", bold=True)

    ref_name = re.sub(r"[^A-Za-z0-9]+", "_", form_data.get("project_manager", "NOM")).strip("_").upper() or "NOM"
    reference = f"E-GES-01_{date.today().strftime('%d%m%y')}_{ref_name}"
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

def render_admin():
    st.title(t("admin_title"))
    admin_password = get_secret("ADMIN_PASSWORD", "")
    if not admin_password:
        st.error(t("not_configured_admin"))
        return
    if not st.session_state.get("admin_authenticated", False):
        with st.form("admin_login_form"):
            password = st.text_input(t("password"), type="password")
            login = st.form_submit_button(t("login"))
        if login:
            if password == str(admin_password):
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.error(t("wrong_password"))
        return

    if st.button(t("logout")):
        st.session_state.admin_authenticated = False
        st.rerun()
    try:
        records = load_metadata()
    except Exception as exc:
        st.error(f"{t('cloud_load_error')} {exc}")
        return

    statuses = [r.get("status", "pending") for r in records]
    st.metric(t("total"), len(records))
    c1, c2, c3 = st.columns(3)
    c1.metric(t("pending"), statuses.count("pending"))
    c2.metric(t("accepted"), statuses.count("accepted"))
    c3.metric(t("refused"), statuses.count("refused"))

    st.subheader(t("by_status"))
    status_labels = {
        "pending": t("status_pending"), "accepted": t("status_accepted"), "refused": t("status_refused")
    }
    counts = {status_labels[k]: statuses.count(k) for k in status_labels}
    st.bar_chart(counts)

    st.subheader(t("by_month"))
    monthly = {}
    for record in records:
        month = str(record.get("submitted_at", ""))[:7]
        if month:
            monthly[month] = monthly.get(month, 0) + 1
    if monthly:
        st.bar_chart(dict(sorted(monthly.items())))
    else:
        st.caption(t("no_requests"))

    st.subheader("Demandes par année / Requests per year")
    yearly = {}
    for record in records:
        year = str(record.get("submitted_at", ""))[:4]
        if year:
            yearly[year] = yearly.get(year, 0) + 1
    if yearly:
        st.bar_chart(dict(sorted(yearly.items())))

    # Export a flat administrative summary without including scientific descriptions.
    import csv
    from io import StringIO
    csv_buffer = StringIO()
    fields = ["id", "project_name", "project_manager", "manager_email", "group",
              "request_date", "submitted_at", "status", "admin_comment", "document_path"]
    writer = csv.DictWriter(csv_buffer, fieldnames=fields, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(records)
    st.download_button(
        "Exporter les statistiques / Export statistics (CSV)",
        data=csv_buffer.getvalue().encode("utf-8-sig"),
        file_name="PO-Chem-WebApp-Request_statistics.csv",
        mime="text/csv",
    )

    st.subheader(t("request_list"))
    if not records:
        st.info(t("no_requests"))
        return
    options = {
        f"{r.get('project_name','(sans nom)')} — {r.get('project_manager','')} — {r.get('submitted_at','')[:10]}": r["id"]
        for r in records
    }
    selected_label = st.selectbox(t("select_request"), list(options.keys()))
    selected_id = options[selected_label]
    record = next(r for r in records if r.get("id") == selected_id)
    st.write(f"**{t('manager')} :** {record.get('project_manager','')}")
    requester_email = record.get("manager_email", "").strip()
    if requester_email:
        st.markdown(f"**{t('manager_email')} :** [{requester_email}](mailto:{requester_email})")
    else:
        st.write(f"**{t('manager_email')} :** —")
    st.write(f"**{t('group')} :** {record.get('group','')}")
    st.write(f"**{t('project_name')} :** {record.get('project_name','')}")
    st.write(f"**{t('doc_path')} :** `{record.get('document_path','')}`")
    status_options = ["pending", "accepted", "refused"]
    status_display = [t("status_pending"), t("status_accepted"), t("status_refused")]
    current_status = record.get("status", "pending")
    current_index = status_options.index(current_status) if current_status in status_options else 0
    with st.form("decision_form"):
        new_status = st.selectbox(t("status"), status_options, index=current_index, format_func=lambda x: status_display[status_options.index(x)])
        comment = st.text_area(t("comment"), value=record.get("admin_comment", ""))
        save = st.form_submit_button(t("save_decision"))
    if save:
        record["status"] = new_status
        record["admin_comment"] = comment.strip()
        record["decision_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
        try:
            save_metadata(records)
        except Exception as exc:
            st.error(f"{t('cloud_error')} {exc}")
            return
        email = record.get("manager_email", "").strip()
        if new_status == "pending":
            st.info("Statut conservé en attente : aucune notification n'est envoyée au demandeur.")
        elif email:
            try:
                send_email(
                    email,
                    t("mail_user_subject"),
                    t("mail_user_body").format(
                        manager=record.get("project_manager", ""),
                        project=record.get("project_name", ""),
                        status=status_display[status_options.index(new_status)],
                        comment=comment.strip() or ("Aucun commentaire." if st.session_state.language == "fr" else "No comment."),
                    ),
                )
                st.success(t("decision_saved"))
            except Exception as exc:
                st.warning(f"{t('decision_saved_no_mail')} {exc}")
        else:
            st.warning(t("select_email_missing"))
        st.rerun()

def render_form():
    if not os.path.isfile(TEMPLATE_PATH):
        st.error("Le modèle Word template_E-GES-01.docx est introuvable.")
        st.stop()
    if "step" not in st.session_state:
        st.session_state.step = 0
    if "data" not in st.session_state:
        st.session_state.data = {}
    data = st.session_state.data
    steps = t("steps")
    step = min(st.session_state.step, len(steps) - 1)

    st.markdown(
        """
        <style>
        .block-container {max-width: 1050px; padding-top: 2rem; padding-bottom: 3rem;}
        .hero {padding: 1.3rem 1.5rem; border-radius: 14px;
               background: linear-gradient(120deg,#edf5fb,#f4f7f2); margin-bottom: 1rem;}
        .hero h1 {margin: 0; color: #183b56; font-size: 1.8rem;}
        .hero p {margin: .4rem 0 0; color: #526575;}
        </style>
        <div class="hero"><h1>OV-Chimie-Métabolisme</h1>
        <p>PO-Chem-WebApp-Request — E-GES-01</p></div>
        """,
        unsafe_allow_html=True,
    )
    st.write(t("intro"))

    def text_field(key, label, **kwargs):
        data[key] = st.text_input(label, value=data.get(key, ""), key=f"field_{key}", **kwargs)

    def text_area(key, label, height=150):
        data[key] = st.text_area(label, value=data.get(key, ""), height=height, key=f"field_{key}")

    st.progress((step + 1) / len(steps), text=f"{t('steps')[step]} — {step + 1}/{len(steps)}")
    with st.container(border=True):
        if step == 0:
            st.subheader(t("general"))
            left, right = st.columns(2)
            with left:
                text_field("project_manager", t("manager"))
                text_field("group", t("group"))
                selected_date = st.date_input(t("date"), value=date.fromisoformat(data["date"]) if data.get("date") else date.today())
                data["date"] = selected_date.isoformat()
                text_field("manager_phone", t("phone"))
                text_field("manager_email", t("manager_email"))
            with right:
                text_field("admin_staff", t("admin_staff"))
                text_field("admin_phone", t("admin_phone"))
                text_field("admin_email", t("admin_email"))
            st.caption(t("recommended"))
        elif step == 1:
            st.subheader(t("description"))
            st.caption(t("description_help"))
            text_area("project_description", t("description_field"), 230)
        elif step == 2:
            st.subheader(t("samples_title"))
            st.caption(t("samples_help"))
            text_area("samples", t("samples_field"), 220)
        elif step == 3:
            st.subheader(t("results_title"))
            text_area("preliminary", t("preliminary"), 160)
            text_area("bibliography", t("bibliography"), 160)
        elif step == 4:
            st.subheader(t("additional"))
            text_field("project_name", t("project_name"))
            yes_no = [t("yes"), t("no")]
            data["funding"] = st.radio(t("funding"), yes_no, horizontal=True, index=0 if data.get("funding") in ("YES", t("yes")) else 1)
            data["funding"] = "YES" if data["funding"] == t("yes") else "NO"
            data["guidance"] = st.radio(t("guidance"), yes_no, horizontal=True, index=0 if data.get("guidance") in ("YES", t("yes")) else 1)
            data["guidance"] = "YES" if data["guidance"] == t("yes") else "NO"
            st.info(t("guidance_info"))
        else:
            st.subheader(t("review"))
            summary = [
                (t("manager"), data.get("project_manager", "")),
                (t("group"), data.get("group", "")),
                (t("date"), data.get("date", "")),
                (t("manager_email"), data.get("manager_email", "")),
                (t("project_name"), data.get("project_name", "")),
                (t("funding"), data.get("funding", "NO")),
                (t("guidance"), data.get("guidance", "NO")),
            ]
            for label, value in summary:
                st.markdown(f"**{label} :** {value or '—'}")
            with st.expander(t("all_answers")):
                for key, value in data.items():
                    st.markdown(f"**{key.replace('_', ' ').capitalize()} :** {value}")
            if not data.get("project_manager") or not data.get("project_name") or not data.get("manager_email"):
                st.warning(t("required"))
            else:
                try:
                    document_bytes, filename = build_document(data)
                    st.download_button(
                        t("download"), data=document_bytes, file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"Erreur de génération Word : {exc}")
                    document_bytes, filename = None, None
                if document_bytes is not None:
                    if st.button(t("submit"), type="primary", use_container_width=True):
                        try:
                            ensure_remote_folder("MyDrive")
                            ensure_remote_folder("MyDrive/E-GES-01")
                            ensure_remote_folder(CLOUD_FOLDER)
                            remote_file = f"{CLOUD_FOLDER}/{filename}"
                            upload = webdav_request("PUT", remote_file, data=document_bytes)
                            upload.raise_for_status()
                            records = load_metadata()
                            request_id = str(uuid.uuid4())
                            record = {
                                "id": request_id,
                                "project_name": data.get("project_name", ""),
                                "project_manager": data.get("project_manager", ""),
                                "manager_email": data.get("manager_email", ""),
                                "group": data.get("group", ""),
                                "request_date": data.get("date", ""),
                                "submitted_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                                "status": "pending",
                                "admin_comment": "",
                                "document_path": remote_file,
                                "filename": filename,
                            }
                            records.append(record)
                            save_metadata(records)
                            st.session_state.last_submission = {
                                "document": document_bytes, "filename": filename, "id": request_id
                            }
                            st.session_state.data = {}
                            st.session_state.step = 0
                            try:
                                # Optional CC for internal follow-up; configured only by the app administrator.
                                admin_cc = str(get_secret("ADMIN_CC_EMAIL", "") or "").strip()
                                send_email(
                                    ADMIN_EMAIL,
                                    t("mail_admin_subject"),
                                    t("mail_admin_body").format(
                                        project=record["project_name"],
                                        manager=record["project_manager"],
                                        group=record["group"],
                                        path=remote_file,
                                    ),
                                    cc=admin_cc or None,
                                    reply_to=record.get("manager_email", "").strip() or None,
                                )
                                st.success(t("submitted"))
                            except Exception as mail_exc:
                                st.warning(f"{t('mail_warning')} {mail_exc}")
                                st.success(t("submitted"))
                        except Exception as exc:
                            st.error(f"{t('cloud_error')} {exc}")

    st.divider()
    if step < len(steps) - 1:
        c1, c2 = st.columns(2)
        with c1:
            if st.button(t("previous"), disabled=(step == 0), use_container_width=True):
                st.session_state.step -= 1
                st.rerun()
        with c2:
            if st.button(t("next"), use_container_width=True):
                st.session_state.step += 1
                st.rerun()
    elif st.button(t("edit"), use_container_width=True):
        st.session_state.step -= 1
        st.rerun()

    if st.session_state.get("last_submission"):
        last = st.session_state.last_submission
        st.download_button(
            t("download"), data=last["document"], file_name=last["filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="download_after_submit",
        )

if "language" not in st.session_state:
    st.session_state.language = "fr"
with st.sidebar:
    st.selectbox(
        TRANSLATIONS[st.session_state.language]["language"],
        options=["fr", "en"],
        format_func=lambda x: "Français" if x == "fr" else "English",
        key="language",
    )
    page = st.radio(t("form"), [t("form"), t("admin")], label_visibility="collapsed")

if page == t("admin"):
    render_admin()
else:
    render_form()
