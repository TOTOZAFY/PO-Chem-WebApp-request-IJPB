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
CLOUD_FOLDER = "MyDrive/Demandes_Analyses"
METADATA_PATH = "MyDrive/demandes_metadata.json"
NOTIFICATION_SETTINGS_PATH = "MyDrive/notification_settings.json"
DEFAULT_ADMIN_EMAIL = "ov-chimie@inrae.fr"
ANALYSIS_TYPES = [
    "Métabolites primaires par GC-MS",
    "Métabolomique non ciblée par LC-HRMS",
    "Analyses d'acides gras par GC-MS",
    "Analyse de stérols totaux par GC-MS",
    "Lipidomique non ciblée LC-HRMS",
    "Lipidomique ciblée LC-MS/MS",
    "Oligoglycomique par LC-HRMS",
    "Monosaccharides par GC-MS",
    "Hormones par LC-MS/MS",
    "Développements spécifiques",
]
SPECIFIC_DEVELOPMENT_TYPE = "Développements spécifiques"
ANALYSIS_TYPES_EN = {
    "Métabolites primaires par GC-MS": "Primary metabolites by GC-MS",
    "Métabolomique non ciblée par LC-HRMS": "Untargeted metabolomics by LC-HRMS",
    "Analyses d'acides gras par GC-MS": "Fatty acid analysis by GC-MS",
    "Analyse de stérols totaux par GC-MS": "Total sterol analysis by GC-MS",
    "Lipidomique non ciblée LC-HRMS": "Untargeted lipidomics by LC-HRMS",
    "Lipidomique ciblée LC-MS/MS": "Targeted lipidomics by LC-MS/MS",
    "Oligoglycomique par LC-HRMS": "Oligoglycomics by LC-HRMS",
    "Monosaccharides par GC-MS": "Monosaccharides by GC-MS",
    "Hormones par LC-MS/MS": "Hormones by LC-MS/MS",
    "Développements spécifiques": "Specific developments",
}
TRANSLATIONS = {
    "fr": {
        "language": "Langue / Language", "form": "Formulaire de demande",
        "admin": "Espace administrateur", "intro": "Complétez les rubriques ci-dessous. "
        "À la dernière étape, votre demande sera enregistrée dans Nextcloud et l'administrateur sera notifié.",
        "steps": ["Contacts", "Description du projet", "Échantillons", "Résultats & bibliographie",
                  "Informations complémentaires", "Vérification & soumission"],
        "general": "1. Informations générales", "requester_section": "Demandeur", "administrator_section": "Plateforme Observatoire du végétal chimie et métabolisme", "manager": "Responsable du projet *",
        "group": "Groupe / laboratoire *", "date": "Date", "phone": "Téléphone",
        "manager_email": "E-mail du responsable *", "admin_staff": "Personnel administratif",
        "admin_phone": "Téléphone du personnel administratif", "admin_email": "E-mail du personnel administratif (facultatif)",
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
        "submitted": "Demande enregistrée dans Nextcloud.",
        "submitted_notified": "Demande enregistrée dans Nextcloud. La notification à l’administrateur a bien été envoyée.",
        "cloud_error": "Le dépôt dans Nextcloud a échoué. La demande n'a pas été enregistrée :",
        "mail_warning": "Le document est enregistré, mais la notification à l'administrateur n'a pas pu être envoyée :",
        "cloud_not_configured": "La connexion Nextcloud n'est pas configurée dans les secrets Streamlit.",
        "admin_login": "Connexion administrateur", "password": "Mot de passe administrateur",
        "login": "Se connecter", "logout": "Se déconnecter", "wrong_password": "Mot de passe incorrect.",
        "admin_title": "Tableau de bord administrateur", "total": "Demandes reçues",
        "pending": "En attente", "accepted": "Acceptées", "refused": "Refusées",
        "by_status": "Répartition par statut", "by_month": "Demandes par mois",
        "request_list": "Gestion des demandes", "no_requests": "Aucune demande enregistrée.",
        "notification_settings": "Configuration des notifications e-mail",
        "admin_primary_email": "E-mail principal de l’administrateur",
        "platform_staff_email": "E-mail du personnel de la plateforme (facultatif)",
        "admin_secondary_email": "E-mail secondaire de l’administrateur (CC, facultatif)",
        "save_notification_settings": "Enregistrer les adresses e-mail",
        "notification_settings_saved": "Adresses de notification enregistrées dans Nextcloud.",
        "select_request": "Sélectionner une demande", "status": "Décision / statut",
        "comment": "Commentaire facultatif", "save_decision": "Enregistrer la décision et notifier le demandeur",
        "decision_saved": "Décision enregistrée et notification envoyée au demandeur.",
        "decision_saved_no_mail": "Décision enregistrée, mais l'e-mail au demandeur n'a pas pu être envoyé :",
        "select_email_missing": "Aucune adresse e-mail du demandeur n'est renseignée.",
        "cloud_load_error": "Impossible de lire les données depuis Nextcloud :",
        "not_configured_admin": "Espace administrateur indisponible : configurez ADMIN_PASSWORD dans les secrets Streamlit.",
        "doc_path": "Document dans Nextcloud", "created": "Soumise le", "status_pending": "En attente",
        "status_accepted": "Acceptée", "status_refused": "Refusée",
        "mail_admin_subject": "Nouvelle demande d'analyse reçue",
        "mail_user_subject": "Mise à jour de votre demande d'analyse",
        "mail_admin_body": "Une nouvelle demande d'analyse a été soumise.\nProjet : {project}\nResponsable : {manager}\nGroupe : {group}\nType d’analyse : {analysis_type}\nDétails : {analysis_details}\nDocument stocké dans Nextcloud : {path}",
        "mail_user_body": "Bonjour {manager},\n\nLa demande concernant le projet « {project} » a été mise à jour.\nDécision : {status}\nCommentaire : {comment}\n\nCordialement,\n",
        "back_form": "Retour au formulaire",
        "platform_contact_email": "E-mail de contact de la plateforme",
        "platform_contact_help": "Cette adresse reçoit les nouvelles demandes. Elle est modifiable dans l'espace de gestion de la plateforme.",
        "platform_contact_caption": "Adresse de notification de la plateforme — affichée ici à titre informatif.",
        "analysis_type": "Type d’analyse *", "analysis_type_stats": "Demandes par type d’analyse",
         "analysis_type_required": "Veuillez sélectionner un type d’analyse.",
        "specific_development_details": "Détails du développement spécifique *",
        "specific_development_help": "Décrivez la demande : métabolites ou molécules ciblés, matrice, objectifs, méthode envisagée, contraintes ou besoins particuliers.",
        "specific_development_required": "Veuillez préciser le développement spécifique demandé.",
        "yearly": "Demandes par année",
        "export_csv": "Exporter les statistiques (CSV)",
        "required_admin_email": "L'adresse e-mail principale de la plateforme est obligatoire.",
        "manage_analysis_types": "Gestion des types d’analyse",
        "add_analysis_type": "Ajouter un type d’analyse",
        "new_analysis_type": "Nom du nouveau type d’analyse",
        "analysis_type_added": "Type d’analyse ajouté.",
        "analysis_type_exists": "Ce type d’analyse existe déjà.",
        "remove_analysis_type": "Supprimer un type d’analyse",
        "analysis_type_removed": "Type d’analyse supprimé.",
        "analysis_type_locked": "Le type « Développements spécifiques » ne peut pas être supprimé car il active le champ de description complémentaire.",
        "save_analysis_type": "Ajouter",
        "delete_analysis_type": "Supprimer",
        "new_analysis_type_en": "Traduction anglaise du type d’analyse",
        "analysis_type_translation_required": "Veuillez renseigner le nom français et la traduction anglaise.",
        "analysis_type_translation_saved": "Type d’analyse et traduction anglaise enregistrés.",
        "edit_analysis_type_translation": "Modifier une traduction anglaise",
        "select_analysis_type_translation": "Type d’analyse à modifier",
        "save_analysis_type_translation": "Enregistrer la traduction",
        "pending_no_notification": "Statut conservé en attente : aucune notification n'est envoyée au demandeur.",
        "template_missing": "Le modèle Word template_E-GES-01.docx est introuvable.",
        "word_generation_error": "Erreur de génération Word :",
        "metadata_invalid": "Le fichier de métadonnées Nextcloud n'est pas un JSON valide.",
        "notification_settings_invalid": "Le fichier de configuration des notifications est invalide.",
        "smtp_not_configured": "SMTP n'est pas configuré dans les secrets Streamlit.",
        "bibliography_prefix": "Bibliographie / Bibliography:",
        "answer_prefix": "Réponse / Answer:",
        "project_manager_short": "Responsable du projet",
        "project_group_short": "Groupe / laboratoire",
        "project_date_short": "Date",
        "project_manager_email_short": "E-mail du responsable",
        "project_staff_short": "Personnel administratif du laboratoire",
        "project_staff_phone_short": "Téléphone du personnel administratif du laboratoire",
        "project_staff_email_short": "E-mail du personnel administratif du laboratoire",
        "platform_staff_email_short": "E-mail du personnel de la plateforme",
        "project_description_short": "Description du projet",
        "samples_short": "Échantillons",
        "preliminary_short": "Résultats préliminaires",
        "bibliography_short": "Bibliographie",
        "project_name_short": "Nom du projet",
        "funding_short": "Financement",
        "guidance_short": "Accompagnement scientifique",
    },
    "en": {
        "language": "Language / Langue", "form": "Request form",
        "admin": "Administrator area", "intro": "Complete the sections below. "
        "At the final step, your request will be saved to Nextcloud and the administrator will be notified.",
        "steps": ["Contacts", "Project description", "Samples", "Preliminary results & bibliography",
                  "Additional information", "Review & submission"],
        "general": "1. General information", "requester_section": "Requester", "administrator_section": "Plateforme Observatoire du végétal chimie et métabolisme", "manager": "Project manager *",
        "group": "Group / laboratory *", "date": "Date", "phone": "Telephone",
        "manager_email": "Project manager email *", "admin_staff": "Administrative staff",
        "admin_phone": "Administrative staff telephone", "admin_email": "Administrative staff email (optional)",
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
        "required": "Required fields: project manager name, email, laboratory / group, and project name. Please also check the email format.",
        "submit": "Submit request", "download": "Download Word document",
        "submitted": "Request saved to Nextcloud.",
        "submitted_notified": "Request saved to Nextcloud. The administrator notification was sent successfully.",
        "cloud_error": "Uploading to Nextcloud failed. The request was not saved:",
        "mail_warning": "The document was saved, but the administrator notification could not be sent:",
        "cloud_not_configured": "Nextcloud connection is not configured in Streamlit secrets.",
        "admin_login": "Administrator login", "password": "Administrator password",
        "login": "Log in", "logout": "Log out", "wrong_password": "Incorrect password.",
        "admin_title": "Administrator dashboard", "total": "Requests received",
        "pending": "Pending", "accepted": "Accepted", "refused": "Rejected",
        "by_status": "Requests by status", "by_month": "Requests per month",
        "request_list": "Request management", "no_requests": "No requests have been recorded.",
        "notification_settings": "Email notification settings",
        "admin_primary_email": "Administrator primary email",
        "platform_staff_email": "Platform staff email (optional)",
        "admin_secondary_email": "Administrator secondary email (CC, optional)",
        "save_notification_settings": "Save email addresses",
        "notification_settings_saved": "Notification addresses saved to Nextcloud.",
        "select_request": "Select a request", "status": "Decision / status",
        "comment": "Optional comment", "save_decision": "Save decision and notify requester",
        "decision_saved": "Decision saved and requester notified.",
        "decision_saved_no_mail": "Decision saved, but the requester email could not be sent:",
        "select_email_missing": "No requester email address is available.",
        "cloud_load_error": "Could not read data from Nextcloud:",
        "not_configured_admin": "Administrator area unavailable: set ADMIN_PASSWORD in Streamlit secrets.",
        "doc_path": "Document in Nextcloud", "created": "Submitted on", "status_pending": "Pending",
        "status_accepted": "Accepted", "status_refused": "Rejected",
        "mail_admin_subject": "New analysis request received",
        "mail_user_subject": "Update on your analysis request",
        "mail_admin_body": "A new analysis request has been submitted.\nProject: {project}\nManager: {manager}\nGroup: {group}\nAnalysis type: {analysis_type}\nDetails: {analysis_details}\nDocument stored in Nextcloud: {path}",
        "mail_user_body": "Hello {manager},\n\nThe request for project “{project}” has been updated.\nDecision: {status}\nComment: {comment}\n\nBest regards,\n",
        "back_form": "Back to request form",
        "platform_contact_email": "Platform contact email",
        "platform_contact_help": "This address receives new requests. It can be changed in the platform management area.",
        "platform_contact_caption": "Platform notification address — displayed here for information.",
        "analysis_type": "Analysis type *", "analysis_type_stats": "Requests by analysis type",
         "analysis_type_required": "Please select an analysis type.",
        "specific_development_details": "Specific development details *",
        "specific_development_help": "Describe the request: target metabolites or molecules, sample matrix, objectives, proposed method, constraints or specific needs.",
        "specific_development_required": "Please describe the requested specific development.",
        "yearly": "Requests per year",
        "export_csv": "Export statistics (CSV)",
        "required_admin_email": "The platform primary email address is required.",
        "manage_analysis_types": "Manage analysis types",
        "add_analysis_type": "Add an analysis type",
        "new_analysis_type": "New analysis type name",
        "analysis_type_added": "Analysis type added.",
        "analysis_type_exists": "This analysis type already exists.",
        "remove_analysis_type": "Remove an analysis type",
        "analysis_type_removed": "Analysis type removed.",
        "analysis_type_locked": "“Specific developments” cannot be removed because it activates the additional details field.",
        "save_analysis_type": "Add",
        "delete_analysis_type": "Remove",
        "new_analysis_type_en": "English translation of the analysis type",
        "analysis_type_translation_required": "Please enter both the French name and the English translation.",
        "analysis_type_translation_saved": "Analysis type and English translation saved.",
        "edit_analysis_type_translation": "Edit an English translation",
        "select_analysis_type_translation": "Analysis type to edit",
        "save_analysis_type_translation": "Save translation",
        "pending_no_notification": "Status remains pending: no notification is sent to the requester.",
        "template_missing": "The Word template template_E-GES-01.docx could not be found.",
        "word_generation_error": "Word generation error:",
        "metadata_invalid": "The Nextcloud metadata file is not valid JSON.",
        "notification_settings_invalid": "The notification settings file is invalid.",
        "smtp_not_configured": "SMTP is not configured in Streamlit secrets.",
        "bibliography_prefix": "Bibliography:",
        "answer_prefix": "Answer:",
        "project_manager_short": "Project manager",
        "project_group_short": "Group / laboratory",
        "project_date_short": "Date",
        "project_manager_email_short": "Project manager email",
        "project_staff_short": "Laboratory administrative staff",
        "project_staff_phone_short": "Laboratory administrative staff telephone",
        "project_staff_email_short": "Laboratory administrative staff email",
        "platform_staff_email_short": "Platform staff email",
        "project_description_short": "Project description",
        "samples_short": "Samples",
        "preliminary_short": "Preliminary results",
        "bibliography_short": "Bibliography",
        "project_name_short": "Project name",
        "funding_short": "Funding",
        "guidance_short": "Scientific guidance",
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
        raise RuntimeError(t("metadata_invalid"))

def save_metadata(records):
    response = webdav_request(
        "PUT", METADATA_PATH,
        data=json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8"),
        timeout=30,
    )
    response.raise_for_status()


def load_notification_settings():
    """Load editable notification recipients from Nextcloud."""
    response = webdav_request("GET", NOTIFICATION_SETTINGS_PATH)
    if response.status_code == 404:
        return {
            "admin_email": DEFAULT_ADMIN_EMAIL,
            "admin_cc_email": "",
            "analysis_types": ANALYSIS_TYPES.copy(),
            "analysis_type_translations": ANALYSIS_TYPES_EN.copy(),
        }
    response.raise_for_status()
    settings = response.json()
    if not isinstance(settings, dict):
        raise RuntimeError(t("notification_settings_invalid"))
    return {
        "admin_email": str(settings.get("admin_email", DEFAULT_ADMIN_EMAIL)).strip(),
        "admin_cc_email": str(settings.get("admin_cc_email", "")).strip(),
        "analysis_types": (
            [str(x).strip() for x in settings.get("analysis_types", ANALYSIS_TYPES)
             if str(x).strip()]
            if isinstance(settings.get("analysis_types", ANALYSIS_TYPES), list)
            else ANALYSIS_TYPES.copy()
        ),
        "analysis_type_translations": {
            **ANALYSIS_TYPES_EN,
            **(
                {str(k).strip(): str(v).strip()
                 for k, v in settings.get("analysis_type_translations", {}).items()
                 if str(k).strip() and str(v).strip()}
                if isinstance(settings.get("analysis_type_translations", {}), dict)
                else {}
            ),
        },
    }


def save_notification_settings(settings):
    response = webdav_request(
        "PUT", NOTIFICATION_SETTINGS_PATH,
        data=json.dumps(settings, ensure_ascii=False, indent=2).encode("utf-8"),
        timeout=30,
    )
    response.raise_for_status()


def join_email_recipients(*addresses):
    """Return a comma-separated list of non-empty, unique email recipients."""
    result = []
    for address in addresses:
        for item in str(address or "").split(","):
            item = item.strip()
            if item and item not in result:
                result.append(item)
    return ", ".join(result) or None

def send_email(to_address, subject, body, cc=None, reply_to=None):
    cfg = get_secret("smtp", {})
    if not cfg or not all(cfg.get(k) for k in ("host", "username", "password", "sender")):
        raise RuntimeError(t("smtp_not_configured"))
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

    # Record the selected analysis type directly in the Word section
    # "Nature et nombre des échantillons", before the sample description.
    samples_cell = table.cell(4, 0)
    analysis_type = form_data.get("analysis_type", "").strip()
    if analysis_type:
        type_paragraph = samples_cell.add_paragraph()
        type_paragraph.paragraph_format.space_after = Pt(4)
        type_run = type_paragraph.add_run(
            ("Type d’analyse : " if st.session_state.get("language", "fr") == "fr" else "Analysis type: ")
            + analysis_type
        )
        type_run.bold = True
        type_run.font.size = Pt(9)

    analysis_details = form_data.get("analysis_details", "").strip()
    if analysis_details:
        details_paragraph = samples_cell.add_paragraph()
        details_paragraph.paragraph_format.space_after = Pt(4)
        details_run = details_paragraph.add_run(
            ("Détails du développement spécifique : " if st.session_state.get("language", "fr") == "fr" else "Specific development details: ")
            + analysis_details
        )
        details_run.font.size = Pt(9)

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
        add_value(results_cell, t("bibliography_prefix") + " " + bibliography)

    project_cell = table.cell(6, 0)
    add_to_existing_paragraph(project_cell.paragraphs[0], form_data.get("project_name"), bold=True)
    add_to_existing_paragraph(project_cell.paragraphs[1], f"{t('answer_prefix')} {form_data.get('funding', 'NO')}", bold=True)
    guidance_cell = table.cell(7, 0)
    add_to_existing_paragraph(guidance_cell.paragraphs[0], f"{t('answer_prefix')} {form_data.get('guidance', 'NO')}", bold=True)

    # Filename format: E-GES-01-Date-Heure-ProjectManager-Lab.docx
    timestamp = datetime.now().astimezone().strftime("%Y%m%d-%H%M%S")
    manager_name = re.sub(
        r"[^A-Za-z0-9]+", "-", form_data.get("project_manager", "NOM")
    ).strip("-") or "NOM"
    lab_name = re.sub(
        r"[^A-Za-z0-9]+", "-", form_data.get("group", "LABO")
    ).strip("-") or "LABO"
    reference = f"E-GES-01-{timestamp}-{manager_name}-{lab_name}"
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

    st.subheader(t("notification_settings"))
    try:
        notification_settings = load_notification_settings()
    except Exception as exc:
        st.error(f"{t('cloud_load_error')} {exc}")
        return

    with st.form("notification_settings_form"):
        admin_primary = st.text_input(
            t("admin_primary_email"),
            value=notification_settings.get("admin_email", DEFAULT_ADMIN_EMAIL),
        )
        admin_secondary = st.text_input(
            t("admin_secondary_email"),
            value=notification_settings.get("admin_cc_email", ""),
        )
        save_recipients = st.form_submit_button(t("save_notification_settings"))
    if save_recipients:
        if not admin_primary.strip():
            st.error(t("required_admin_email"))
        else:
            try:
                save_notification_settings({
                    "admin_email": admin_primary.strip(),
                    "admin_cc_email": admin_secondary.strip(),
                    "analysis_types": notification_settings.get("analysis_types", ANALYSIS_TYPES),
                    "analysis_type_translations": notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN),
                })
                st.success(t("notification_settings_saved"))
                st.rerun()
            except Exception as exc:
                st.error(f"{t('cloud_error')} {exc}")

    st.subheader(t("manage_analysis_types"))
    managed_types = notification_settings.get("analysis_types") or ANALYSIS_TYPES.copy()
    with st.form("add_analysis_type_form"):
        new_type = st.text_input(t("new_analysis_type"))
        new_type_en = st.text_input(t("new_analysis_type_en"))
        add_type = st.form_submit_button(t("save_analysis_type"))
    if add_type:
        candidate = new_type.strip()
        candidate_en = new_type_en.strip()
        if not candidate or not candidate_en:
            st.warning(t("analysis_type_translation_required"))
        elif any(candidate.casefold() == item.casefold() for item in managed_types):
            st.warning(t("analysis_type_exists"))
        else:
            try:
                managed_types.append(candidate)
                type_translations = notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN.copy())
                type_translations[candidate] = candidate_en
                save_notification_settings({
                    "admin_email": notification_settings.get("admin_email", DEFAULT_ADMIN_EMAIL),
                    "admin_cc_email": notification_settings.get("admin_cc_email", ""),
                    "analysis_types": managed_types,
                    "analysis_type_translations": type_translations,
                })
                st.success(t("analysis_type_translation_saved"))
                st.rerun()
            except Exception as exc:
                st.error(f"{t('cloud_error')} {exc}")

    st.markdown(f"**{t('edit_analysis_type_translation')}**")
    current_type_translations = notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN.copy())
    with st.form("edit_analysis_type_translation_form"):
        type_to_translate = st.selectbox(
            t("select_analysis_type_translation"),
            options=managed_types,
            key="type_to_translate",
        )
        edited_translation = st.text_input(
            t("new_analysis_type_en"),
            value=current_type_translations.get(type_to_translate, ""),
            key="edited_analysis_type_translation",
        )
        save_translation = st.form_submit_button(t("save_analysis_type_translation"))
    if save_translation:
        if not edited_translation.strip():
            st.warning(t("analysis_type_translation_required"))
        else:
            try:
                current_type_translations[type_to_translate] = edited_translation.strip()
                save_notification_settings({
                    "admin_email": notification_settings.get("admin_email", DEFAULT_ADMIN_EMAIL),
                    "admin_cc_email": notification_settings.get("admin_cc_email", ""),
                    "analysis_types": managed_types,
                    "analysis_type_translations": current_type_translations,
                })
                st.success(t("analysis_type_translation_saved"))
                st.rerun()
            except Exception as exc:
                st.error(f"{t('cloud_error')} {exc}")

    removable_types = [item for item in managed_types if item != SPECIFIC_DEVELOPMENT_TYPE]
    if removable_types:
        with st.form("remove_analysis_type_form"):
            type_to_remove = st.selectbox(t("remove_analysis_type"), options=removable_types)
            remove_type = st.form_submit_button(t("delete_analysis_type"))
        if remove_type:
            try:
                updated_types = [item for item in managed_types if item != type_to_remove]
                if not updated_types:
                    st.warning(t("analysis_type_required"))
                else:
                    save_notification_settings({
                        "admin_email": notification_settings.get("admin_email", DEFAULT_ADMIN_EMAIL),
                        "admin_cc_email": notification_settings.get("admin_cc_email", ""),
                        "analysis_types": updated_types,
                        "analysis_type_translations": notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN),
                    })
                    st.success(t("analysis_type_removed"))
                    st.rerun()
            except Exception as exc:
                st.error(f"{t('cloud_error')} {exc}")
    st.caption(t("analysis_type_locked"))

    # Archived requests remain in the same Nextcloud metadata file; only their
    # archived flag changes. This preserves existing records and compatibility.
    active_records = [r for r in records if not r.get("archived", False)]
    archived_records = [r for r in records if r.get("archived", False)]

    active_tab, archive_tab = st.tabs([
        "📋 Demandes actives" if st.session_state.language == "fr" else "📋 Active requests",
        "🗄️ Archives" if st.session_state.language == "fr" else "🗄️ Archive",
    ])

    with active_tab:
        statuses = [r.get("status", "pending") for r in active_records]
        st.metric(t("total"), len(active_records))
        c1, c2, c3 = st.columns(3)
        c1.metric(t("pending"), statuses.count("pending"))
        c2.metric(t("accepted"), statuses.count("accepted"))
        c3.metric(t("refused"), statuses.count("refused"))

        st.subheader(t("analysis_type_stats"))
        analysis_counts = {}
        for item in active_records:
            analysis_type = item.get("analysis_type") or ("Non renseigné" if st.session_state.language == "fr" else "Not specified")
            display_type = analysis_type if st.session_state.language == "fr" else notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN).get(analysis_type, ANALYSIS_TYPES_EN.get(analysis_type, analysis_type))
            analysis_counts[display_type] = analysis_counts.get(display_type, 0) + 1
        if analysis_counts:
            st.bar_chart(dict(sorted(analysis_counts.items())))
        else:
            st.caption(t("no_requests"))

        st.subheader(t("by_status"))
        status_labels = {
            "pending": t("status_pending"), "accepted": t("status_accepted"), "refused": t("status_refused")
        }
        counts = {status_labels[k]: statuses.count(k) for k in status_labels}
        st.bar_chart(counts)

        st.subheader(t("by_month"))
        monthly = {}
        for item in active_records:
            month = str(item.get("submitted_at", ""))[:7]
            if month:
                monthly[month] = monthly.get(month, 0) + 1
        if monthly:
            st.bar_chart(dict(sorted(monthly.items())))
        else:
            st.caption(t("no_requests"))

        st.subheader(t("yearly"))
        yearly = {}
        for item in active_records:
            year = str(item.get("submitted_at", ""))[:4]
            if year:
                yearly[year] = yearly.get(year, 0) + 1
        if yearly:
            st.bar_chart(dict(sorted(yearly.items())))

        import csv
        from io import StringIO
        csv_buffer = StringIO()
        fields = ["id", "project_name", "project_manager", "manager_email", "group",
                  "request_date", "submitted_at", "status", "admin_comment", "document_path", "quote_signed", "signed_quote_path", "analysis_type", "analysis_details"]
        writer = csv.DictWriter(csv_buffer, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(active_records)
        st.download_button(
            t("export_csv"),
            data=csv_buffer.getvalue().encode("utf-8-sig"),
            file_name="PO-Chem-WebApp-Request_statistics.csv",
            mime="text/csv",
        )

        st.subheader(t("request_list"))
        if not active_records:
            st.info(t("no_requests"))
        else:
            options = {
                f"{r.get('project_name','(sans nom)')} — {r.get('project_manager','')} — {str(r.get('submitted_at',''))[:10]}": r["id"]
                for r in active_records
            }
            selected_label = st.selectbox(t("select_request"), list(options.keys()), key="active_request_select")
            selected_id = options[selected_label]
            record = next(r for r in active_records if r.get("id") == selected_id)
            st.write(f"**{t('manager')} :** {record.get('project_manager','')}")
            requester_email = record.get("manager_email", "").strip()
            if requester_email:
                st.markdown(f"**{t('manager_email')} :** [{requester_email}](mailto:{requester_email})")
            else:
                st.write(f"**{t('manager_email')} :** —")
            st.write(f"**{t('group')} :** {record.get('group','')}")
            st.write(f"**{t('project_name')} :** {record.get('project_name','')}")
            stored_type = record.get("analysis_type", "")
            shown_type = stored_type if st.session_state.language == "fr" else notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN).get(stored_type, ANALYSIS_TYPES_EN.get(stored_type, stored_type))
            st.write(f"**{t('analysis_type')} :** {shown_type or '—'}")
            if record.get("analysis_details"):
                details_label = "Détails du développement spécifique" if st.session_state.language == "fr" else "Specific development details"
                st.write(f"**{details_label} :** {record.get('analysis_details')}")
            st.write(f"**{t('doc_path')} :** `{record.get('document_path','')}`")

            # Signed quotation management: the file is stored separately from request documents.
            quote_signed_label = "Devis signé" if st.session_state.language == "fr" else "Quotation signed"
            quote_upload_label = "Importer le devis signé (PDF, DOC ou DOCX)" if st.session_state.language == "fr" else "Upload signed quotation (PDF, DOC or DOCX)"
            quote_save_label = "Enregistrer le statut et le devis" if st.session_state.language == "fr" else "Save status and quotation"
            existing_quote_path = record.get("signed_quote_path", "")
            with st.form(f"signed_quote_form_{selected_id}"):
                quote_signed = st.checkbox(
                    quote_signed_label,
                    value=bool(record.get("quote_signed", False)),
                )
                quote_file = st.file_uploader(
                    quote_upload_label,
                    type=["pdf", "doc", "docx"],
                    key=f"signed_quote_upload_{selected_id}",
                )
                if existing_quote_path:
                    st.caption(
                        ("Devis déjà déposé : " if st.session_state.language == "fr" else "Quotation already uploaded: ")
                        + existing_quote_path
                    )
                save_quote = st.form_submit_button(quote_save_label)
            if save_quote:
                if quote_signed and not (quote_file or existing_quote_path):
                    st.warning(
                        "Coche le statut uniquement après avoir importé le devis signé."
                        if st.session_state.language == "fr"
                        else "Upload the signed quotation before marking it as signed."
                    )
                else:
                    try:
                        new_quote_path = existing_quote_path
                        if quote_file is not None:
                            ensure_remote_folder("MyDrive")
                            ensure_remote_folder("MyDrive/Devis_Signes")
                            safe_original = re.sub(r"[^A-Za-z0-9._-]+", "-", os.path.basename(quote_file.name)).strip("-")
                            safe_original = safe_original or "devis-signe"
                            new_quote_path = f"MyDrive/Devis_Signes/{selected_id}_{safe_original}"
                            quote_response = webdav_request("PUT", new_quote_path, data=quote_file.getvalue())
                            quote_response.raise_for_status()
                        record["quote_signed"] = bool(quote_signed)
                        record["signed_quote_path"] = new_quote_path
                        record["signed_quote_filename"] = quote_file.name if quote_file is not None else record.get("signed_quote_filename", "")
                        record["quote_updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
                        save_metadata(records)
                        st.success(
                            "Statut du devis enregistré dans Nextcloud."
                            if st.session_state.language == "fr"
                            else "Quotation status saved to Nextcloud."
                        )
                        st.rerun()
                    except Exception as exc:
                        st.error(
                            f"Impossible d'enregistrer le devis : {exc}"
                            if st.session_state.language == "fr"
                            else f"Could not save quotation: {exc}"
                        )
            status_options = ["pending", "accepted", "refused"]
            status_display = [t("status_pending"), t("status_accepted"), t("status_refused")]
            current_status = record.get("status", "pending")
            current_index = status_options.index(current_status) if current_status in status_options else 0
            with st.form("decision_form"):
                new_status = st.selectbox(t("status"), status_options, index=current_index,
                    format_func=lambda x: status_display[status_options.index(x)])
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
                    st.info(t("pending_no_notification"))
                elif email:
                    try:
                        send_email(
                            email, t("mail_user_subject"),
                            t("mail_user_body").format(
                                manager=record.get("project_manager", ""),
                                project=record.get("project_name", ""),
                                status=status_display[status_options.index(new_status)],
                                comment=comment.strip() or ("Aucun commentaire." if st.session_state.language == "fr" else "No comment."),
                            ),
                            cc=join_email_recipients(record.get("admin_email", ""), record.get("platform_staff_email", "")),
                        )
                        st.success(t("decision_saved"))
                    except Exception as exc:
                        st.warning(f"{t('decision_saved_no_mail')} {exc}")
                else:
                    st.warning(t("select_email_missing"))
                st.rerun()

            st.divider()
            archive_label = "Archiver cette demande" if st.session_state.language == "fr" else "Archive this request"
            if st.button("🗄️ " + archive_label, key=f"archive_{selected_id}"):
                record["archived"] = True
                record["archived_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
                try:
                    save_metadata(records)
                    st.success("Demande déplacée dans les archives." if st.session_state.language == "fr" else "Request moved to archive.")
                    st.rerun()
                except Exception as exc:
                    record.pop("archived", None)
                    record.pop("archived_at", None)
                    st.error(f"{t('cloud_error')} {exc}")

    with archive_tab:
        st.metric("Demandes archivées" if st.session_state.language == "fr" else "Archived requests", len(archived_records))
        if not archived_records:
            st.info("Aucune demande archivée." if st.session_state.language == "fr" else "No archived requests.")
        else:
            archive_options = {
                f"{r.get('project_name','(sans nom)')} — {r.get('project_manager','')} — {str(r.get('submitted_at',''))[:10]}": r["id"]
                for r in archived_records
            }
            archive_label = st.selectbox(
                "Sélectionner une demande archivée" if st.session_state.language == "fr" else "Select an archived request",
                list(archive_options.keys()), key="archived_request_select"
            )
            archived_id = archive_options[archive_label]
            archived_record = next(r for r in archived_records if r.get("id") == archived_id)
            st.write(f"**{t('manager')} :** {archived_record.get('project_manager','')}")
            archived_email = archived_record.get("manager_email", "").strip()
            if archived_email:
                st.markdown(f"**{t('manager_email')} :** [{archived_email}](mailto:{archived_email})")
            st.write(f"**{t('group')} :** {archived_record.get('group','')}")
            st.write(f"**{t('project_name')} :** {archived_record.get('project_name','')}")
            st.write(f"**{t('status')} :** {archived_record.get('status','pending')}")
            st.write(f"**Archivée le :** {archived_record.get('archived_at','—')}" if st.session_state.language == "fr" else f"**Archived on:** {archived_record.get('archived_at','—')}")
            st.write(f"**{t('doc_path')} :** `{archived_record.get('document_path','')}`")

            restore_label = "Restaurer dans les demandes actives" if st.session_state.language == "fr" else "Restore to active requests"
            if st.button("↩️ " + restore_label, key=f"restore_{archived_id}"):
                archived_record.pop("archived", None)
                archived_record.pop("archived_at", None)
                try:
                    save_metadata(records)
                    st.success("Demande restaurée." if st.session_state.language == "fr" else "Request restored.")
                    st.rerun()
                except Exception as exc:
                    archived_record["archived"] = True
                    st.error(f"{t('cloud_error')} {exc}")

            st.divider()
            confirm_label = "Je confirme la suppression définitive" if st.session_state.language == "fr" else "I confirm permanent deletion"
            delete_label = "Supprimer définitivement" if st.session_state.language == "fr" else "Delete permanently"
            confirmed = st.checkbox(confirm_label, key=f"confirm_delete_{archived_id}")
            if st.button("🗑️ " + delete_label, type="primary", disabled=not confirmed, key=f"delete_{archived_id}"):
                document_path = archived_record.get("document_path", "")
                try:
                    # Delete the associated Word document first. Keep the metadata
                    # untouched if Nextcloud reports a real deletion error.
                    if document_path:
                        delete_response = webdav_request("DELETE", document_path)
                        if delete_response.status_code not in (200, 202, 204, 404):
                            delete_response.raise_for_status()
                    records = [r for r in records if r.get("id") != archived_id]
                    save_metadata(records)
                    st.success("Demande et document supprimés définitivement." if st.session_state.language == "fr" else "Request and document permanently deleted.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Suppression impossible : {exc}" if st.session_state.language == "fr" else f"Deletion failed: {exc}")

def render_form():
    if not os.path.isfile(TEMPLATE_PATH):
        st.error(t("template_missing"))
        st.stop()
    if "step" not in st.session_state:
        st.session_state.step = 0
    if "data" not in st.session_state:
        st.session_state.data = {}
    data = st.session_state.data
    try:
        form_analysis_settings = load_notification_settings()
        available_analysis_types = form_analysis_settings.get("analysis_types") or ANALYSIS_TYPES.copy()
        available_analysis_type_translations = form_analysis_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN)
    except Exception:
        available_analysis_types = ANALYSIS_TYPES.copy()
        available_analysis_type_translations = ANALYSIS_TYPES_EN.copy()
    if SPECIFIC_DEVELOPMENT_TYPE not in available_analysis_types:
        available_analysis_types.append(SPECIFIC_DEVELOPMENT_TYPE)
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
        <div class="hero"><h1>OV-Chimie-Métabolisme/IJPB-INRAE Versailles</h1>
        <p>PO-Chem-WebApp-Request</p></div>
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
            st.markdown(f"**{t('requester_section')}**")
            requester_left, requester_right = st.columns(2)
            with requester_left:
                text_field("project_manager", t("manager"))
                text_field("group", t("group"))
                selected_date = st.date_input(t("date"), value=date.fromisoformat(data["date"]) if data.get("date") else date.today())
                data["date"] = selected_date.isoformat()
                text_field("manager_phone", t("phone"))
                text_field("manager_email", t("manager_email"))
            with requester_right:
                text_field("admin_staff", t("admin_staff"))
                text_field("admin_phone", t("admin_phone"))
                text_field("admin_email", t("admin_email"))

            st.divider()
            st.markdown(f"**{t('administrator_section')}**")
            try:
                current_admin_email = load_notification_settings().get("admin_email", DEFAULT_ADMIN_EMAIL)
            except Exception:
                current_admin_email = DEFAULT_ADMIN_EMAIL
            st.text_input(
                t("platform_contact_email"),
                value=current_admin_email,
                key="platform_admin_email_display",
                disabled=True,
                help=t("platform_contact_help"),
            )
            st.caption(
                t("platform_contact_caption")
            )
            current_analysis_type = data.get("analysis_type", available_analysis_types[0])
            analysis_index = available_analysis_types.index(current_analysis_type) if current_analysis_type in available_analysis_types else 0
            analysis_labels = available_analysis_types if st.session_state.language == "fr" else [
                available_analysis_type_translations.get(item, ANALYSIS_TYPES_EN.get(item, item)) for item in available_analysis_types
            ]
            selected_analysis_label = st.selectbox(
                t("analysis_type"),
                options=analysis_labels,
                index=analysis_index,
                key="analysis_type_select",
            )
            data["analysis_type"] = (
                selected_analysis_label if st.session_state.language == "fr"
                else available_analysis_types[analysis_labels.index(selected_analysis_label)]
            )
            if data["analysis_type"] == SPECIFIC_DEVELOPMENT_TYPE:
                st.text_area(
                    t("specific_development_details"),
                    value=data.get("analysis_details", ""),
                    height=130,
                    help=t("specific_development_help"),
                    key="analysis_details_input",
                )
                data["analysis_details"] = st.session_state.get("analysis_details_input", "")
            else:
                data["analysis_details"] = ""
            text_field("platform_staff_email", t("platform_staff_email"))
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
                (t("analysis_type"), data.get("analysis_type", "")),
                (t("specific_development_details"), data.get("analysis_details", "") if data.get("analysis_type") == SPECIFIC_DEVELOPMENT_TYPE else ""),
                (t("project_name"), data.get("project_name", "")),
                (t("funding"), data.get("funding", "NO")),
                (t("guidance"), data.get("guidance", "NO")),
            ]
            for label, value in summary:
                st.markdown(f"**{label} :** {value or '—'}")
            with st.expander(t("all_answers")):
                answer_labels = {
                    "project_manager": t("project_manager_short"),
                    "group": t("project_group_short"),
                    "date": t("project_date_short"),
                    "manager_phone": t("phone"),
                    "manager_email": t("project_manager_email_short"),
                    "admin_staff": t("project_staff_short"),
                    "admin_phone": t("project_staff_phone_short"),
                    "admin_email": t("project_staff_email_short"),
                    "platform_staff_email": t("platform_staff_email_short"),
                    "analysis_type": t("analysis_type"),
                    "analysis_details": t("specific_development_details"),
                    "project_description": t("project_description_short"),
                    "samples": t("samples_short"),
                    "preliminary": t("preliminary_short"),
                    "bibliography": t("bibliography_short"),
                    "project_name": t("project_name_short"),
                    "funding": t("funding_short"),
                    "guidance": t("guidance_short"),
                }
                for key, value in data.items():
                    st.markdown(f"**{answer_labels.get(key, key.replace('_', ' ').capitalize())} :** {value}")
            required_values = [
                data.get("project_manager", "").strip(),
                data.get("manager_email", "").strip(),
                data.get("group", "").strip(),
                data.get("project_name", "").strip(),
                data.get("analysis_type", "").strip(),
            ]
            email_is_valid = bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", data.get("manager_email", "").strip()))
            specific_details_valid = (
                data.get("analysis_type") != SPECIFIC_DEVELOPMENT_TYPE
                or bool(data.get("analysis_details", "").strip())
            )
            form_is_valid = all(required_values) and email_is_valid and specific_details_valid
            if not form_is_valid:
                if data.get("analysis_type") == SPECIFIC_DEVELOPMENT_TYPE and not data.get("analysis_details", "").strip():
                    st.warning(t("specific_development_required"))
                else:
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
                    st.error(f"{t('word_generation_error')} {exc}")
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
                                "admin_staff": data.get("admin_staff", ""),
                                "admin_phone": data.get("admin_phone", ""),
                                "admin_email": data.get("admin_email", ""),
                                "platform_staff_email": data.get("platform_staff_email", ""),
                                "analysis_type": data.get("analysis_type", ""),
                                "analysis_details": data.get("analysis_details", ""),
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
                                # Recipients are editable in the administrator area and stored in Nextcloud.
                                recipients = load_notification_settings()
                                admin_email = recipients.get("admin_email") or DEFAULT_ADMIN_EMAIL
                                internal_cc = join_email_recipients(
                                    recipients.get("admin_cc_email"),
                                    record.get("admin_email", ""),
                                    record.get("platform_staff_email", ""),
                                )
                                send_email(
                                    admin_email,
                                    t("mail_admin_subject"),
                                    t("mail_admin_body").format(
                                        project=record["project_name"],
                                        manager=record["project_manager"],
                                        group=record["group"],
                                        analysis_type=record.get("analysis_type", ""),
                                        analysis_details=record.get("analysis_details", "") or ("Non applicable" if st.session_state.language == "fr" else "Not applicable"),
                                        path=remote_file,
                                    ),
                                    cc=internal_cc,
                                    reply_to=record.get("manager_email", "").strip() or None,
                                )
                                st.success(t("submitted_notified"))
                            except Exception as mail_exc:
                                st.warning(f"{t('mail_warning')} {mail_exc}")
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
