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
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
import csv

st.set_page_config(
    page_title="PO-Chem-WebApp-Request | OV-Chimie-Métabolisme",
    page_icon="🧪",
    layout="wide",
)

APP_NAME = "PO-Chem-WebApp-Request"
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_E-GES-01.docx")
FAIR_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "FAIR_template.xlsx")
CLOUD_FOLDER = "MyDrive/Demandes_Analyses"
METADATA_PATH = "MyDrive/demandes_metadata.json"
NOTIFICATION_SETTINGS_PATH = "MyDrive/notification_settings.json"
USER_PROFILES_PATH = "MyDrive/user_profiles.json"
FAIR_FOLDER = CLOUD_FOLDER
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


# Persistent interface profiles. Only interface preferences are stored here;
# request contents, FAIR values, credentials and form text are deliberately excluded.
PROFILE_FIXED_PREFERENCES = {
    # Preferences specific to FAIR > Analysis only. Global navigation and
    # application language are intentionally not profile-specific.
    "fair_analysis_role", "analysis_type_select",
}


def load_user_profiles():
    response = webdav_request("GET", USER_PROFILES_PATH)
    if response.status_code == 404:
        return {"profiles": {}}
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict) or not isinstance(payload.get("profiles", {}), dict):
        raise RuntimeError("Le fichier des profils utilisateur dans Nextcloud est invalide.")
    return payload


def save_user_profiles(payload):
    response = webdav_request(
        "PUT", USER_PROFILES_PATH,
        data=json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        timeout=30,
    )
    response.raise_for_status()


def _profile_sheet_keys(sheet):
    """Return the session-state keys belonging to one FAIR sheet."""
    if sheet == "Analysis":
        return set(PROFILE_FIXED_PREFERENCES), "analysis_protocol_"
    if sheet == "Data treatment":
        return set(), "fair_data_treatment_pref_"
    if sheet == "Pipeline Step by Step":
        return set(), "fair_pipeline_step_pref_"
    return set(), ""


def get_profile_preferences(sheet="Analysis"):
    fixed_keys, prefix = _profile_sheet_keys(sheet)
    preferences = {}
    for key in fixed_keys:
        if key in st.session_state:
            value = st.session_state[key]
            if isinstance(value, (str, int, float, bool)) or value is None:
                preferences[key] = value
    if prefix:
        for key, value in st.session_state.items():
            if key.startswith(prefix) and not key.endswith("_text"):
                if isinstance(value, (str, int, float, bool)) or value is None:
                    preferences[key] = value
    return preferences


def apply_profile_preferences(preferences, sheet="Analysis"):
    fixed_keys, prefix = _profile_sheet_keys(sheet)
    for key in list(st.session_state.keys()):
        if key in fixed_keys or (prefix and key.startswith(prefix) and not key.endswith("_text")):
            st.session_state.pop(key, None)
    for key, value in preferences.items():
        if key in fixed_keys or (prefix and key.startswith(prefix) and not key.endswith("_text")):
            st.session_state[key] = value


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

    # Filename format includes the unique request ID.
    request_id = form_data.get("request_id", "REQ-SANS-ID")
    timestamp = datetime.now().astimezone().strftime("%Y%m%d-%H%M%S")
    manager_name = re.sub(
        r"[^A-Za-z0-9]+", "-", form_data.get("project_manager", "NOM")
    ).strip("-") or "NOM"
    lab_name = re.sub(
        r"[^A-Za-z0-9]+", "-", form_data.get("group", "LABO")
    ).strip("-") or "LABO"
    reference = f"E-GES-01-{request_id}-{timestamp}-{manager_name}-{lab_name}"
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

def fair_filename(record):
    """Keep the request document stem and append the FAIR suffix."""
    stem = os.path.splitext(record.get("filename", record.get("id", "request")))[0]
    return f"{stem}_FAIR.xlsx"


def fair_remote_path(record):
    return f"{FAIR_FOLDER}/{fair_filename(record)}"


def create_fair_workbook(record):
    if not os.path.isfile(FAIR_TEMPLATE_PATH):
        raise FileNotFoundError("FAIR_template.xlsx is missing")
    wb = load_workbook(FAIR_TEMPLATE_PATH)
    ws = wb["Sample"]
    # Clear example values while retaining the original layout, labels and styles.
    for row in range(1, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):
            if ws.cell(row, 1).value is not None and not isinstance(ws.cell(row, col), MergedCell):
                ws.cell(row, col).value = None
    ws["B1"] = record.get("project_name", "")
    ws["B2"] = (record.get("submitted_at", "") or record.get("request_date", ""))[:10]
    ws["B3"] = record.get("id", "")
    ws["B4"] = record.get("analysis_type", "")
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def load_fair_workbook(record):
    path = record.get("fair_path") or fair_remote_path(record)
    response = webdav_request("GET", path)
    if response.status_code == 404:
        payload = create_fair_workbook(record)
        upload = webdav_request("PUT", path, data=payload)
        upload.raise_for_status()
        record["fair_path"] = path
        return load_workbook(BytesIO(payload))
    response.raise_for_status()
    return load_workbook(BytesIO(response.content))


def save_fair_workbook(record, workbook):
    output = BytesIO()
    workbook.save(output)
    path = record.get("fair_path") or fair_remote_path(record)
    response = webdav_request("PUT", path, data=output.getvalue())
    response.raise_for_status()
    record["fair_path"] = path
    record["fair_updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")


def sample_rows(ws):
    return {str(ws.cell(r, 1).value).strip(): r for r in range(1, ws.max_row + 1)
            if ws.cell(r, 1).value is not None and str(ws.cell(r, 1).value).strip()}


def fair_merge_samples(ws, incoming, requester=True):
    """Merge sample dictionaries by tube ID; never replace admin-owned fields from requester imports."""
    rows = sample_rows(ws)
    id_row = rows.get("ID tube analyse")
    if not id_row:
        raise ValueError("The Sample sheet has no 'ID tube analyse' row.")
    admin_only = {"File ID Q-Tof ESI+", "File ID Q-Tof ESI-"}
    existing = {}
    for col in range(2, ws.max_column + 1):
        tube = ws.cell(id_row, col).value
        if tube is not None and str(tube).strip():
            existing[str(tube).strip()] = col
    seen = set()
    for sample in incoming:
        tube = str(sample.get("ID tube analyse", "")).strip()
        if not tube:
            raise ValueError("Each sample must have an ID tube analyse.")
        if tube in seen:
            raise ValueError(f"Duplicate tube ID in imported data: {tube}")
        seen.add(tube)
        col = existing.get(tube)
        if col is None:
            col = next((candidate for candidate in range(2, ws.max_column + 1)
                        if ws.cell(id_row, candidate).value in (None, "")), ws.max_column + 1)
            existing[tube] = col
        for field, value in sample.items():
            field = str(field).strip()
            if not field or value is None or str(value).strip() == "":
                continue
            if requester and field in admin_only:
                continue
            row = rows.get(field)
            if row is None:
                row = ws.max_row + 1
                ws.cell(row, 1).value = field
                rows[field] = row
            ws.cell(row, col).value = value
    return len(incoming)


def parse_sample_upload(upload):
    name = upload.name.lower()
    if name.endswith((".xlsx", ".xlsm")):
        wb = load_workbook(BytesIO(upload.getvalue()), data_only=True)
        ws = wb[wb.sheetnames[0]]
        # Accepted layout: metadata names in column A, samples in columns B+.
        labels = [str(ws.cell(r, 1).value).strip() if ws.cell(r, 1).value is not None else ""
                  for r in range(1, ws.max_row + 1)]
        id_index = next((i for i, x in enumerate(labels, 1) if x == "ID tube analyse"), None)
        if id_index is None:
            # Also accept a conventional table with headers in row 1.
            headers = [str(ws.cell(1, c).value).strip() if ws.cell(1, c).value is not None else ""
                       for c in range(1, ws.max_column + 1)]
            if "ID tube analyse" not in headers:
                raise ValueError("Missing 'ID tube analyse' field.")
            return [{headers[c-1]: ws.cell(r, c).value for c in range(1, len(headers)+1) if headers[c-1]}
                    for r in range(2, ws.max_row+1) if any(ws.cell(r,c).value is not None for c in range(1,len(headers)+1))]
        samples=[]
        aliases = {"File ID Q-Tof ESI−": "File ID Q-Tof ESI-"}
        for col in range(2, ws.max_column+1):
            item={}
            for row, label in enumerate(labels, 1):
                if label:
                    item[aliases.get(label, label)]=ws.cell(row,col).value
            if any(v is not None and str(v).strip() for v in item.values()): samples.append(item)
        return samples
    # CSV/TSV in conventional row-per-sample format.
    text_data = upload.getvalue().decode("utf-8-sig")
    dialect = csv.Sniffer().sniff(text_data[:4096], delimiters=";,\t")
    reader = csv.DictReader(text_data.splitlines(), dialect=dialect)
    aliases = {"File ID Q-Tof ESI−": "File ID Q-Tof ESI-"}
    return [{aliases.get(k, k): v for k, v in row.items() if k} for row in reader if any(v and v.strip() for v in row.values())]


def render_fair_requester(record, records):
    fr = st.session_state.get("language", "fr") == "fr"
    title = "FAIR – Métadonnées des échantillons" if fr else "FAIR – Sample metadata"
    st.title(title)
    st.caption(("Demande : " if fr else "Request: ") + record.get("id", ""))
    try:
        wb = load_fair_workbook(record)
        ws = wb["Sample"]
        rows = sample_rows(ws)
        # Requester-owned general field and sample metadata form.
        st.subheader("1. " + ("Informations générales" if fr else "General information"))
        current_collection = ws.cell(rows.get("Plant Collection", 7), 2).value or ""
        collection = st.text_area("Plant Collection", value=str(current_collection), height=130)
        st.subheader("2. " + ("Métadonnées des échantillons" if fr else "Sample metadata"))
        warning = ("Le modèle de métadonnées est fourni à titre d'exemple. Veuillez adapter les champs à votre expérience. Vous pouvez ajouter des lignes pour toute information expérimentale supplémentaire et laisser vides les rubriques non applicables. Conservez les intitulés standards utilisés pour la correspondance automatique." if fr else "The provided metadata template is an example. Please adapt the metadata fields to your own experiment: keep the fields that are relevant, add rows for any additional experimental information, and leave non-applicable fields blank. Do not change the standard field names where they are used for automatic matching.")
        st.info(warning)
        mode = st.radio("Mode de saisie" if fr else "Entry method", [
            "Formulaire web" if fr else "Web form",
            "Import Excel/CSV", "Tableau interactif" if fr else "Interactive table",
            "Télécharger le modèle FAIR" if fr else "Download FAIR template"], horizontal=False)
        if mode in ("Formulaire web", "Web form"):
            n = st.number_input("Nombre d'échantillons à saisir" if fr else "Number of samples", min_value=1, max_value=500, value=1, step=1)
            for i in range(int(n)):
                with st.expander(("Échantillon " if fr else "Sample ") + str(i+1), expanded=(i==0)):
                    vals={}
                    for field in rows:
                        if field in {"Metadata", "File ID Q-Tof ESI+", "File ID Q-Tof ESI-", "Name of the experiment", "Date of the request", "N° saved file", "Type of analysis", "Plant Collection"}:
                            continue
                        default=""
                        vals[field]=st.text_input(field, key=f"fair_form_{record['id']}_{i}_{field}", value=default)
                    if "ID tube analyse" not in vals:
                        vals["ID tube analyse"]=st.text_input("ID tube analyse", key=f"fair_form_id_{record['id']}_{i}")
                    if st.button(("Ajouter cet échantillon" if fr else "Add this sample"), key=f"fair_add_{record['id']}_{i}"):
                        try:
                            fair_merge_samples(ws,[vals],requester=True)
                            ws.cell(rows.get("Plant Collection",7),2).value=collection
                            save_fair_workbook(record,wb); save_metadata(records)
                            st.success("Échantillon enregistré." if fr else "Sample saved.")
                        except Exception as exc: st.error(str(exc))
        elif mode == "Import Excel/CSV":
            upload=st.file_uploader("Importer Excel ou CSV", type=["xlsx","xlsm","csv","tsv"], key=f"fair_import_{record['id']}")
            if upload and st.button("Importer les métadonnées" if fr else "Import metadata", key=f"fair_import_btn_{record['id']}"):
                try:
                    incoming=parse_sample_upload(upload); count=fair_merge_samples(ws,incoming,requester=True)
                    ws.cell(rows.get("Plant Collection",7),2).value=collection
                    save_fair_workbook(record,wb); save_metadata(records)
                    st.success((f"{count} échantillon(s) importé(s)." if fr else f"{count} sample(s) imported."))
                except Exception as exc: st.error(str(exc))
        elif mode in ("Tableau interactif", "Interactive table"):
            labels=list(rows.keys()); idrow=rows.get("ID tube analyse",11)
            data=[]
            for col in range(2,ws.max_column+1):
                tube=ws.cell(idrow,col).value
                if tube is not None:
                    data.append({field:ws.cell(r,col).value for field,r in rows.items() if field not in {"File ID Q-Tof ESI+","File ID Q-Tof ESI-"}})
            import pandas as pd
            frame=pd.DataFrame(data, columns=[x for x in labels if x not in {"File ID Q-Tof ESI+","File ID Q-Tof ESI-","Metadata"}])
            edited=st.data_editor(frame, num_rows="dynamic", use_container_width=True, key=f"fair_editor_{record['id']}")
            if st.button("Enregistrer le tableau" if fr else "Save table", key=f"fair_editor_save_{record['id']}"):
                try:
                    # Rebuild only requester-owned sample fields, retaining admin File IDs by tube ID.
                    fair_merge_samples(ws,edited.fillna("").to_dict(orient="records"),requester=True)
                    ws.cell(rows.get("Plant Collection",7),2).value=collection
                    save_fair_workbook(record,wb); save_metadata(records)
                    st.success("Métadonnées enregistrées." if fr else "Metadata saved.")
                except Exception as exc: st.error(str(exc))
        else:
            output=BytesIO(); wb.save(output)
            st.download_button("Télécharger le modèle FAIR" if fr else "Download FAIR template", output.getvalue(), file_name=fair_filename(record), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.divider()
        if st.button("Enregistrer Plant Collection" if fr else "Save Plant Collection", key=f"fair_collection_{record['id']}"):
            ws.cell(rows.get("Plant Collection",7),2).value=collection
            save_fair_workbook(record,wb); save_metadata(records)
            st.success("Plant Collection enregistrée." if fr else "Plant Collection saved.")
    except Exception as exc:
        st.error(("Impossible de charger le fichier FAIR : " if fr else "Could not load FAIR workbook: ") + str(exc))



def render_fair_administrator(record, records):
    """Administrator-only FAIR fields. Requester-owned metadata are never displayed here."""
    fr = st.session_state.get("language", "fr") == "fr"
    st.title("FAIR – Espace administrateur" if fr else "FAIR – Administrator area")
    st.caption(("Demande : " if fr else "Request: ") + str(record.get("id", "")))
    try:
        wb = load_fair_workbook(record)
        ws = wb["Sample"]
        rows = sample_rows(ws)
        id_row = rows.get("ID tube analyse")
        plus_row = rows.get("File ID Q-Tof ESI+")
        minus_row = rows.get("File ID Q-Tof ESI-")
        if not id_row or not plus_row or not minus_row:
            st.error("Les lignes ID tube analyse / File ID sont introuvables dans Sample." if fr else "ID tube analyse / File ID rows are missing from Sample.")
            return
        st.subheader("Sample – Informations réservées à l'administrateur" if fr else "Sample – Administrator-only information")
        st.info("Seuls les identifiants de fichiers instrumentaux sont affichés ici. Les métadonnées saisies par le demandeur restent masquées et sont conservées." if fr else "Only instrument file identifiers are shown here. Requester metadata remain hidden and are preserved.")
        records_table = []
        for col in range(2, ws.max_column + 1):
            tube = ws.cell(id_row, col).value
            if tube not in (None, ""):
                records_table.append({
                    "ID tube analyse": str(tube),
                    "File ID Q-Tof ESI+": ws.cell(plus_row, col).value or "",
                    "File ID Q-Tof ESI-": ws.cell(minus_row, col).value or "",
                })
        import pandas as pd
        frame = pd.DataFrame(records_table, columns=["ID tube analyse", "File ID Q-Tof ESI+", "File ID Q-Tof ESI-"])
        edited = st.data_editor(frame, use_container_width=True, hide_index=True, disabled=["ID tube analyse"], key=f"fair_admin_editor_{record.get('id')}")
        if st.button("Enregistrer les File IDs" if fr else "Save File IDs", key=f"fair_admin_editor_save_{record.get('id')}"):
            try:
                fair_merge_samples(ws, edited.fillna("").to_dict(orient="records"), requester=False)
                save_fair_workbook(record, wb)
                save_metadata(records)
                st.success("File IDs enregistrés." if fr else "File IDs saved.")
            except Exception as exc:
                st.error(str(exc))
        with st.expander("Importer les File IDs depuis Excel/CSV" if fr else "Import File IDs from Excel/CSV"):
            upload = st.file_uploader("Excel / CSV", type=["xlsx", "xlsm", "csv", "tsv"], key=f"fair_admin_import_{record.get('id')}")
            if upload and st.button("Importer les File IDs" if fr else "Import File IDs", key=f"fair_admin_import_btn_{record.get('id')}"):
                try:
                    incoming = parse_sample_upload(upload)
                    existing = {str(ws.cell(id_row, c).value).strip() for c in range(2, ws.max_column + 1) if ws.cell(id_row, c).value not in (None, "")}
                    unknown = [str(item.get("ID tube analyse", "")).strip() for item in incoming if str(item.get("ID tube analyse", "")).strip() not in existing]
                    if unknown:
                        raise ValueError(("ID tube analyse inconnu(s) : " if fr else "Unknown sample ID(s): ") + ", ".join(unknown))
                    admin_rows = [{k: v for k, v in item.items() if k in {"ID tube analyse", "File ID Q-Tof ESI+", "File ID Q-Tof ESI-"}} for item in incoming]
                    fair_merge_samples(ws, admin_rows, requester=False)
                    save_fair_workbook(record, wb)
                    save_metadata(records)
                    st.success("File IDs importés." if fr else "File IDs imported.")
                except Exception as exc:
                    st.error(str(exc))
    except Exception as exc:
        st.error(("Impossible de charger le fichier FAIR : " if fr else "Could not load FAIR workbook: ") + str(exc))



# FAIR Analysis protocol presets. Texts are kept editable and saved in the original Analysis sheet.
FAIR_EXTRACTION_PROTOCOLS = {
    "Protocole de base – Routaboul et al. (2006)": "The metabolite extraction protocol was adapted from Routaboul et al. (2006). Briefly, 1.8 mL of extraction solution (methanol/water/acetone/formic acid, proportion: 30/28/42/0.8 % v/v) and 500 ng of Apigenin (used as internal standard) were added to X mg lyophilized powder. Each sample was homogenized using a FastPrep instrument (2 × 30 sec, 14,000 rpm). The samples were placed in an ultrasonic bath (Advantage-Lab, AL04-04-230) for 1 min at 25 kHz, shaken for 30 min at 4°C using a ThermoMixer C (Eppendorf), and centrifuged for 10 min at maximum speed to remove debris. The supernatants were transferred to new tubes and the pellets were extracted again with 1.8 mL of extraction solution, ultra-sonicated, shaken and centrifuged as in the first step. The second supernatant was pooled with the first one in a glass tube, and dried in a SpeedVac rotary evaporator and freeze dryer. The dry pellet was resuspended in 200 µL of ULC/MS grade water (Biosolve) with 10% acetonitrile ULC/MS grade (Biosolve) and filtered through Whatman glass filter (grade GF/D, 2.7 µm pore size).",
    "Proto Massi modifié": "Extraction: Metabolites were extracted from X mg of camelina dry mature seeds. Briefly, 1 mL of MeOH: Methyl-tert-butyl: H2O (1:3:1), conserved at 4°C, and 200 ng of Apigenin (used as internal standard) were added to each sample, which were then homogenized in 2 mL tubes using a FastPrep instrument (1 min, 14,000 rpm). The mixtures were then shaken for 30 min at 4°C using a ThermoMixer C (Eppendorf), placed in an ice-cooled ultrasonication bath for 15 min and centrifuged for 1 min at maximum speed to remove debris. The extracted samples were then transferred to new tubes containing 650 mL of MeOH:water (1:3), previously placed at -20°C. The mixtures were centrifuged for 1 min at 14,000 rpm. The addition of MeOH:water (1:3) and centrifugation led to phase separation, providing the upper organic phase containing lipids, a lower aqueous phase containing polar and semi-polar metabolites, and a pellet of starch and proteins. The phase containing polar and semi-polar metabolites was dried down in a SpeedVac vacuum concentrator (overnight) and resuspended in X µL of ULC/MS grade water (Biosolve) with 10% ACN ULC/MS grade (Biosolve), then filtered through a Whatman filter.",
    "Proto Jean Marc modifié": "The metabolite extraction protocol was adapted from Routaboul et al. (2006). Briefly, 1.8 mL of extraction solution (methanol/water/acetone/formic acid, proportion: 30/28/42/0.8 % v/v) and 500 ng of Apigenin (used as internal standard) were added to X mg lyophilized powder. Each sample was homogenized using a FastPrep instrument (2 × 30 sec, 14,000 rpm). The samples were placed in an ultrasonic bath (Advantage-Lab, AL04-04-230) for 1 min at 25 kHz, shaken for 30 min at 4°C using a ThermoMixer C (Eppendorf), and centrifuged for 10 min at maximum speed to remove debris. The supernatants were transferred to new tubes and the pellets were extracted again with 1.8 mL of extraction solution, ultra-sonicated, shaken and centrifuged as in the first step. The second supernatant was pooled with the first one in a glass tube, and dried in a SpeedVac rotary evaporator and freeze dryer. The dry pellet was resuspended in 200 µL of ULC/MS grade water (Biosolve) with 10% acetonitrile ULC/MS grade (Biosolve) and filtered through Whatman glass filter (grade GF/D, 2.7 µm pore size).",
    "Proto Gilles": "All steps were performed in 2 mL Safelock Eppendorf tubes. The ground frozen samples (50 mg) were resuspended in 1 mL of frozen (-20°C) Water:Acetonitrile:Isopropanol (2:3:3) containing 200 ng apigenin (used as internal standard) and extracted for 10 min at 4°C with shaking at 1400 rpm in an Eppendorf Thermomixer. Insoluble material was removed by centrifugation at 20000g for 5 min. The supernatant was collected and dried in a Speed-Vac.",
}

def render_fair_analysis(record, records):
    fr = st.session_state.get("language", "fr") == "fr"
    st.title("FAIR – Analysis")
    st.caption(("Demande : " if fr else "Request: ") + str(record.get("id", "")))
    st.info("Chaque champ doit être vérifié et renseigné selon la manipulation réellement effectuée. Les textes proposés sont des exemples modifiables : adaptez-les, supprimez les éléments non applicables et précisez toute information manquante." if fr else "Check and complete every field according to the experiment actually performed. Suggested texts are editable examples: adapt them, remove non-applicable information and specify any missing details.")
    try:
        wb = load_fair_workbook(record)
        ws = wb["Analysis"]
        # The FAIR template stores field names in column A and values in column B.
        labels = {str(ws.cell(r, 1).value).strip(): r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value is not None and str(ws.cell(r, 1).value).strip()}
        protocol_label = "Extraction protocol"
        if protocol_label not in labels:
            st.error("Le champ Extraction protocol est introuvable dans la feuille Analysis." if fr else "Extraction protocol field is missing from the Analysis sheet.")
            return
        protocol_row = labels[protocol_label]
        saved_text = str(ws.cell(protocol_row, 2).value or "")
        preset_names = list(FAIR_EXTRACTION_PROTOCOLS.keys()) + (["Autre / Personnaliser"] if fr else ["Other / Customize"])
        # Try to identify the current preset from the saved content.
        selected_default = "Autre / Personnaliser" if fr else "Other / Customize"
        for name, body in FAIR_EXTRACTION_PROTOCOLS.items():
            if saved_text.strip() == body.strip():
                selected_default = name
                break
        st.subheader("1. " + ("Protocole d'extraction" if fr else "Extraction protocol"))
        selected_protocol = st.selectbox("Choisir un protocole" if fr else "Select a protocol", preset_names, index=preset_names.index(selected_default), key=f"analysis_protocol_{record.get('id')}")
        if selected_protocol in FAIR_EXTRACTION_PROTOCOLS:
            proposed = FAIR_EXTRACTION_PROTOCOLS[selected_protocol]
        else:
            proposed = saved_text
        protocol_text = st.text_area("Description complète du protocole – modifiable" if fr else "Full protocol description – editable", value=proposed, height=300, key=f"analysis_protocol_text_{record.get('id')}_{selected_protocol}")
        if selected_protocol == ("Autre / Personnaliser" if fr else "Other / Customize") and not protocol_text.strip():
            st.caption("Veuillez décrire le protocole réellement utilisé." if fr else "Please describe the protocol actually used.")
        st.divider()
        st.subheader("2. " + ("Autres informations Analysis" if fr else "Other Analysis information"))
        # Existing FAIR labels and example values are retained; every field is editable by either profile.
        excluded = {protocol_label}
        field_values = {}
        for field, row in labels.items():
            if field in excluded or field == "Acquisition":
                continue
            current = ws.cell(row, 2).value
            if current is None:
                current = ""
            field_values[field] = st.text_area(field, value=str(current), height=100 if len(str(current)) > 120 else 68, key=f"analysis_field_{record.get('id')}_{row}")
        if st.button("Soumettre et enregistrer Analysis" if fr else "Submit and save Analysis", type="primary", key=f"analysis_save_{record.get('id')}"):
            ws.cell(protocol_row, 2).value = protocol_text.strip()
            # The selected protocol is stored only in the FAIR "Extraction protocol" field (column B).
            # Do not duplicate preset names or protocol texts in auxiliary columns.
            ws.cell(protocol_row, 3).value = None
            for obsolete_cell in ("D2", "E2", "D10", "E10", "D11", "E11"):
                ws[obsolete_cell].value = None
            for field, value in field_values.items():
                ws.cell(labels[field], 2).value = value.strip()
            save_fair_workbook(record, wb)
            save_metadata(records)
            st.success("Les informations de la feuille Analysis ont été enregistrées dans le classeur FAIR." if fr else "Analysis information has been saved to the FAIR workbook.")
    except Exception as exc:
        st.error(("Impossible de charger ou d'enregistrer la feuille Analysis : " if fr else "Could not load or save the Analysis sheet: ") + str(exc))

def render_fair_portal():
    """FAIR navigation: choose a sheet first, then a role within that sheet."""
    fr = st.session_state.get("language", "fr") == "fr"
    st.title("FAIR")

    # First level: FAIR workbook sheets.
    sheet_options = ["Sample", "Analysis", "Data treatment", "Pipeline Step by Step"]
    selected_sheet = st.selectbox(
        "Feuille FAIR" if fr else "FAIR sheet",
        options=sheet_options,
        key="fair_selected_sheet",
    )

    if selected_sheet == "Analysis":
        st.divider()
        render_fair_profile_manager("Analysis")
        st.divider()
        st.subheader("Analysis")
        role = st.selectbox(
            "Qui souhaite compléter la feuille Analysis ?" if fr else "Who is completing the Analysis sheet?",
            options=["Demandeur", "Administrateur"] if fr else ["Requester", "Administrator"],
            key="fair_analysis_role",
        )
        records = load_metadata()
        active_records = [r for r in records if not r.get("deleted", False)]
        requester_role = "Demandeur" if fr else "Requester"
        if role == requester_role:
            with st.form("fair_analysis_requester_access_form"):
                reference = st.text_input("Référence de la demande" if fr else "Request reference")
                email = st.text_input("Adresse e-mail du demandeur" if fr else "Requester email address")
                submitted = st.form_submit_button("Accéder à Analysis" if fr else "Open Analysis")
            if submitted:
                record = next((r for r in active_records if str(r.get("id", "")).strip().casefold() == reference.strip().casefold() and str(r.get("manager_email", "")).strip().casefold() == email.strip().casefold()), None)
                if not record:
                    st.error("Référence ou adresse e-mail incorrecte." if fr else "Request reference or email address is incorrect.")
                    return
                st.session_state["fair_analysis_requester_id"] = record.get("id")
                st.rerun()
            saved_id = st.session_state.get("fair_analysis_requester_id")
            record = next((r for r in active_records if str(r.get("id")) == str(saved_id)), None) if saved_id else None
            if record:
                render_fair_analysis(record, records)
            return
        admin_password = get_secret("ADMIN_PASSWORD", "")
        if not admin_password:
            st.error("Le secret ADMIN_PASSWORD n'est pas configuré." if fr else "ADMIN_PASSWORD secret is not configured.")
            return
        if not st.session_state.get("fair_analysis_admin_authenticated", False):
            with st.form("fair_analysis_admin_login_form"):
                password = st.text_input("Mot de passe administrateur" if fr else "Administrator password", type="password")
                login = st.form_submit_button("Se connecter" if fr else "Log in")
            if login and password == admin_password:
                st.session_state["fair_analysis_admin_authenticated"] = True
                st.rerun()
            elif login:
                st.error("Mot de passe incorrect." if fr else "Incorrect password.")
            return
        if active_records:
            choices = {f"{r.get('id')} – {r.get('project_name', '')}": r for r in active_records}
            selected = st.selectbox("Sélectionner une demande" if fr else "Select a request", list(choices), key="fair_analysis_admin_request")
            render_fair_analysis(choices[selected], records)
        else:
            st.info("Aucune demande disponible." if fr else "No requests available.")
        return

    if selected_sheet in ("Data treatment", "Pipeline Step by Step"):
        st.divider()
        render_fair_profile_manager(selected_sheet)
        st.divider()
        st.subheader(selected_sheet)
        st.info(
            "La feuille sera développée dans une prochaine étape. Les préférences propres à cette feuille sont déjà mémorisées dans le profil actif."
            if fr else
            "This sheet will be implemented in a later phase. Sheet-specific preferences are already stored in the active profile."
        )
        return

    st.divider()
    st.subheader("Sample")
    # Second level: choose who is completing the Sample sheet.
    role = st.selectbox(
        "Qui souhaite compléter la feuille Sample ?" if fr else "Who is completing the Sample sheet?",
        options=["Demandeur", "Administrateur"] if fr else ["Requester", "Administrator"],
        key="fair_sample_role",
    )

    records = load_metadata()
    active_records = [r for r in records if not r.get("deleted", False)]

    requester_role = "Demandeur" if fr else "Requester"
    if role == requester_role:
        st.write(
            "Saisissez la référence de votre demande et l'adresse e-mail utilisée lors de sa création. Vous accéderez uniquement aux champs à compléter par le demandeur."
            if fr else
            "Enter your request reference and the email address used when submitting it. You will only access fields assigned to the requester."
        )
        with st.form("fair_sample_requester_access_form"):
            reference = st.text_input("Référence de la demande" if fr else "Request reference")
            email = st.text_input("Adresse e-mail du demandeur" if fr else "Requester email address")
            submitted = st.form_submit_button("Accéder à Sample" if fr else "Open Sample")
        if submitted:
            norm_ref, norm_email = reference.strip().casefold(), email.strip().casefold()
            record = next((r for r in active_records
                           if str(r.get("id", "")).strip().casefold() == norm_ref
                           and str(r.get("manager_email", "")).strip().casefold() == norm_email), None)
            if not record:
                st.error("Référence ou adresse e-mail incorrecte. Vérifiez les informations indiquées." if fr else "Reference or email address not recognized. Please check your details.")
                return
            if record.get("fair_status") not in ("requested", "in_progress", "completed"):
                st.warning("Le remplissage FAIR n'a pas encore été demandé par l'administrateur." if fr else "The administrator has not requested FAIR completion yet.")
                return
            st.session_state["fair_sample_requester_id"] = record.get("id")
            st.rerun()

        saved_id = st.session_state.get("fair_sample_requester_id")
        record = next((r for r in active_records if str(r.get("id")) == str(saved_id)), None) if saved_id else None
        if record and record.get("fair_status") in ("requested", "in_progress", "completed"):
            st.divider()
            render_fair_requester(record, records)
        return

    # Administrator authentication is scoped to the Sample administrator area.
    st.write(
        "L'administrateur peut compléter uniquement les champs qui lui sont réservés dans Sample."
        if fr else
        "The administrator can complete only the fields assigned to them in Sample."
    )
    admin_password = get_secret("ADMIN_PASSWORD", "")
    if not admin_password:
        st.error("Le secret ADMIN_PASSWORD n'est pas configuré." if fr else "ADMIN_PASSWORD secret is not configured.")
        return
    if not st.session_state.get("fair_sample_admin_authenticated", False):
        with st.form("fair_sample_admin_login_form"):
            password = st.text_input("Mot de passe administrateur" if fr else "Administrator password", type="password")
            login = st.form_submit_button("Se connecter" if fr else "Log in")
        if login:
            if password == str(admin_password):
                st.session_state["fair_sample_admin_authenticated"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect." if fr else "Incorrect password.")
        return

    if st.button("Se déconnecter de Sample administrateur" if fr else "Log out of Sample administrator"):
        st.session_state["fair_sample_admin_authenticated"] = False
        st.session_state.pop("fair_sample_admin_request_id", None)
        st.rerun()

    if not active_records:
        st.info("Aucune demande disponible." if fr else "No requests available.")
        return
    labels = {f"{r.get('id', '')} — {r.get('project_name', '')} — {r.get('project_manager', '')}": r for r in active_records}
    selected = st.selectbox("Sélectionner une demande" if fr else "Select a request", list(labels), key="fair_sample_admin_request_select")
    record = labels[selected]
    st.session_state["fair_sample_admin_request_id"] = record.get("id")
    render_fair_administrator(record, records)


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

            st.subheader("FAIR – Sample")
            fair_path = record.get("fair_path") or fair_remote_path(record)
            st.caption(("Classeur FAIR : " if st.session_state.language == "fr" else "FAIR workbook: ") + fair_path)
            fair_col1, fair_col2 = st.columns(2)
            with fair_col1:
                if st.button("Créer / vérifier le classeur FAIR" if st.session_state.language == "fr" else "Create / check FAIR workbook", key=f"fair_create_{selected_id}"):
                    try:
                        fair_wb = load_fair_workbook(record); save_fair_workbook(record, fair_wb); save_metadata(records)
                        st.success("Classeur FAIR disponible." if st.session_state.language == "fr" else "FAIR workbook is available.")
                    except Exception as exc: st.error(str(exc))
            with fair_col2:
                if st.button("Demander le remplissage FAIR" if st.session_state.language == "fr" else "Request FAIR completion", key=f"fair_request_{selected_id}"):
                    try:
                        if not record.get("fair_token"): record["fair_token"] = uuid.uuid4().hex
                        app_url = str(get_secret("APP_BASE_URL", "")).strip().rstrip("/")
                        if not app_url:
                            st.warning("Ajoutez APP_BASE_URL dans les secrets Streamlit pour générer le lien sécurisé." if st.session_state.language == "fr" else "Add APP_BASE_URL to Streamlit secrets to generate the secure link.")
                        else:
                            link = f"{app_url}/?fair_token={record['fair_token']}"
                            subject = "Compléter les métadonnées FAIR – " + record.get("project_name", "")
                            if st.session_state.language == "fr":
                                body = (
                                    "Bonjour " + record.get("project_manager", "") + ",\n\n"
                                    "Merci de compléter les métadonnées FAIR de votre demande.\n"
                                    "Référence de votre demande : " + str(record.get("id", "")) + "\n"
                                    "Vous pouvez ouvrir l'application, choisir le menu FAIR, puis saisir cette référence et votre adresse e-mail de demandeur.\n\n"
                                    "Accès direct sécurisé :\n" + link
                                )
                            else:
                                body = (
                                    "Hello " + record.get("project_manager", "") + ",\n\n"
                                    "Please complete the FAIR metadata for your request.\n"
                                    "Your request reference: " + str(record.get("id", "")) + "\n"
                                    "Open the application, select the FAIR menu, then enter this reference and your requester email address.\n\n"
                                    "Secure direct access: \n" + link
                                )
                            send_email(record.get("manager_email", ""), subject, body)
                            record["fair_status"] = "requested"; record["fair_requested_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
                            save_metadata(records)
                            st.success("Demande FAIR envoyée." if st.session_state.language == "fr" else "FAIR completion request sent.")
                    except Exception as exc: st.error(str(exc))
            try:
                fair_wb = load_fair_workbook(record); fair_ws = fair_wb["Sample"]; fair_rows = sample_rows(fair_ws)
                tube_row = fair_rows.get("ID tube analyse", 11)
                sample_count = sum(1 for c in range(2, fair_ws.max_column+1) if fair_ws.cell(tube_row,c).value not in (None,""))
                st.caption((f"{sample_count} échantillon(s) enregistré(s)" if st.session_state.language == "fr" else f"{sample_count} sample(s) recorded"))
                with st.expander("Importer les File IDs administrateur" if st.session_state.language == "fr" else "Import administrator File IDs"):
                    admin_upload = st.file_uploader("Excel / CSV", type=["xlsx","xlsm","csv","tsv"], key=f"fair_admin_upload_{selected_id}")
                    if admin_upload and st.button("Mettre à jour les File IDs" if st.session_state.language == "fr" else "Update File IDs", key=f"fair_admin_save_{selected_id}"):
                        try:
                            admin_samples = parse_sample_upload(admin_upload)
                            existing_ids = {str(fair_ws.cell(tube_row,c).value).strip() for c in range(2,fair_ws.max_column+1) if fair_ws.cell(tube_row,c).value not in (None,"")}
                            unknown = [str(x.get("ID tube analyse", "")).strip() for x in admin_samples if str(x.get("ID tube analyse", "")).strip() not in existing_ids]
                            if unknown: raise ValueError("ID tube inconnu(s) : " + ", ".join(unknown))
                            admin_only_rows = [{k: v for k, v in item.items() if k in {"ID tube analyse", "File ID Q-Tof ESI+", "File ID Q-Tof ESI-"}} for item in admin_samples]
                            fair_merge_samples(fair_ws, admin_only_rows, requester=False); save_fair_workbook(record, fair_wb); save_metadata(records)
                            st.success("File IDs mis à jour." if st.session_state.language == "fr" else "File IDs updated.")
                        except Exception as exc: st.error(str(exc))
            except Exception as exc: st.warning(str(exc))

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
                    fair_path = archived_record.get("fair_path") or fair_remote_path(archived_record)
                    fair_delete = webdav_request("DELETE", fair_path)
                    if fair_delete.status_code not in (200, 202, 204, 404): fair_delete.raise_for_status()
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
                if not data.get("request_id"):
                    data["request_id"] = (
                        f"REQ-{datetime.now().astimezone().year}-{uuid.uuid4().hex[:8].upper()}"
                    )
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
                            request_id = data["request_id"]
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
                            fair_bytes = create_fair_workbook(record)
                            fair_path = fair_remote_path(record)
                            fair_upload = webdav_request("PUT", fair_path, data=fair_bytes)
                            fair_upload.raise_for_status()
                            record["fair_path"] = fair_path
                            record["fair_status"] = "created"
                            record["fair_token"] = uuid.uuid4().hex
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

if "language" not in st.session_state:
    st.session_state.language = "fr"

# Profiles are loaded and displayed only inside FAIR > Analysis.
profile_payload = {"profiles": {}}
profile_storage_ok = False
profile_names = []

def render_fair_profile_manager(sheet):
    """Show and activate a persistent profile within selected FAIR sheets only."""
    global profile_payload, profile_storage_ok, profile_names
    fr = st.session_state.get("language", "fr") == "fr"
    try:
        profile_payload = load_user_profiles()
        profile_storage_ok = True
        if not profile_payload["profiles"]:
            profile_payload["profiles"]["Chriso"] = {"preferences_by_sheet": {}}
            save_user_profiles(profile_payload)
    except Exception as profile_exc:
        profile_storage_ok = False
        st.warning(("Profils indisponibles : " if fr else "Profiles unavailable: ") + str(profile_exc))

    profile_names = sorted(profile_payload.get("profiles", {}).keys(), key=str.casefold)
    if not profile_names:
        profile_names = ["Session temporaire"]
    if "active_user_profile" not in st.session_state or st.session_state.active_user_profile not in profile_names:
        st.session_state.active_user_profile = "Chriso" if "Chriso" in profile_names else profile_names[0]

    profile_data = profile_payload.get("profiles", {}).get(st.session_state.active_user_profile, {})
    # Migrate the former v1.7-v1.9 single-sheet structure without losing preferences.
    prefs_by_sheet = profile_data.setdefault("preferences_by_sheet", {})
    if "preferences" in profile_data and "Analysis" not in prefs_by_sheet:
        prefs_by_sheet["Analysis"] = profile_data.pop("preferences")
    selected_prefs = prefs_by_sheet.get(sheet, {})
    loaded_marker = f"_loaded_user_profile_{sheet}"
    if st.session_state.get(loaded_marker) != st.session_state.active_user_profile:
        apply_profile_preferences(selected_prefs, sheet)
        st.session_state[loaded_marker] = st.session_state.active_user_profile

    st.subheader("Profil utilisateur" if fr else "User profile")
    chosen_profile = st.selectbox(
        "Profil actif" if fr else "Active profile",
        options=profile_names,
        index=profile_names.index(st.session_state.active_user_profile),
        key=f"fair_{sheet.lower().replace(' ', '_')}_active_profile_select",
    )
    if chosen_profile != st.session_state.active_user_profile:
        st.session_state.active_user_profile = chosen_profile
        st.session_state.pop(loaded_marker, None)
        st.rerun()
    st.caption(
        f"Le profil mémorise uniquement les préférences de l'interface FAIR → {sheet}. Les réponses et métadonnées restent dans le classeur FAIR de chaque demande."
        if fr else
        f"The profile stores only FAIR → {sheet} interface preferences. Answers and metadata remain in each request's FAIR workbook."
    )
    with st.expander("Ajouter un autre profil" if fr else "Add another profile"):
        st.caption(
            "Chaque profil conserve séparément ses préférences pour Analysis, Data Treatment et Pipeline Step by Step."
            if fr else
            "Each profile keeps separate preferences for Analysis, Data Treatment and Pipeline Step by Step."
        )
        with st.form(f"create_user_profile_form_{sheet.lower().replace(' ', '_')}", clear_on_submit=True):
            new_profile_name = st.text_input(
                "Nom du nouveau profil" if fr else "New profile name",
                placeholder="Ex. Marie, Jean, Chriso…" if fr else "e.g. Marie, Jean, Chris…",
            )
            create_profile_submitted = st.form_submit_button(
                "Ajouter le profil" if fr else "Add profile", use_container_width=True
            )
        if create_profile_submitted:
            clean_name = new_profile_name.strip()
            if not clean_name:
                st.warning("Saisis un nom de profil." if fr else "Enter a profile name.")
            elif any(existing.casefold() == clean_name.casefold() for existing in profile_names):
                st.warning("Ce profil existe déjà. Choisis un autre nom." if fr else "This profile already exists. Choose another name.")
            elif not profile_storage_ok:
                st.error("Le stockage Nextcloud des profils n'est pas disponible." if fr else "Nextcloud profile storage is unavailable.")
            else:
                profile_payload["profiles"][clean_name] = {
                    "preferences_by_sheet": {sheet: get_profile_preferences(sheet)}
                }
                try:
                    save_user_profiles(profile_payload)
                    st.session_state.active_user_profile = clean_name
                    st.session_state.pop(loaded_marker, None)
                    st.rerun()
                except Exception as exc:
                    profile_payload["profiles"].pop(clean_name, None)
                    st.error(("Impossible de créer le profil : " if fr else "Could not create profile: ") + str(exc))

# Main navigation. Keep the internal page key explicit so the routing below
# never references an undefined `page` variable.
if st.query_params.get("fair_token"):
    page = "fair_link"
else:
    _fr_navigation = st.session_state.get("language", "fr") == "fr"
    _navigation_options = {
        ("Demande d'analyse" if _fr_navigation else "Analysis request"): "form",
        ("FAIR" if _fr_navigation else "FAIR"): "fair",
        ("Administrateur" if _fr_navigation else "Administrator"): "admin",
    }
    _selected_navigation_label = st.sidebar.radio(
        "Navigation",
        options=list(_navigation_options.keys()),
        index=list(_navigation_options.values()).index(
            st.session_state.get("main_navigation_page", "form")
            if st.session_state.get("main_navigation_page", "form") in _navigation_options.values()
            else "form"
        ),
        key="main_navigation_selector",
    )
    page = _navigation_options[_selected_navigation_label]
    st.session_state["main_navigation_page"] = page

if st.query_params.get("fair_token"):
    try:
        fair_records = load_metadata()
        token = st.query_params.get("fair_token")
        fair_record = next((r for r in fair_records if r.get("fair_token") == token), None)
        if not fair_record:
            st.error("Lien FAIR invalide ou expiré." if st.session_state.language == "fr" else "Invalid or expired FAIR link.")
        elif fair_record.get("fair_status") not in ("requested", "in_progress", "completed"):
            st.warning("Le remplissage FAIR n'a pas encore été demandé par l'administrateur." if st.session_state.language == "fr" else "FAIR completion has not yet been requested by the administrator.")
        else:
            fair_record["fair_status"] = "in_progress"
            render_fair_requester(fair_record, fair_records)
    except Exception as exc:
        st.error(f"FAIR: {exc}")
elif page == "admin":
    render_admin()
elif page == "fair":
    render_fair_portal()
else:
    render_form()

# Persist preferences only inside the three selected FAIR sheets with profile support.
# Sample and all other application sections remain outside the profile system.
_selected_profile_sheet = st.session_state.get("fair_selected_sheet")
if (page == "fair" and _selected_profile_sheet in ("Analysis", "Data treatment", "Pipeline Step by Step")
        and profile_storage_ok
        and st.session_state.get("active_user_profile") in profile_payload.get("profiles", {})):
    try:
        _profile = profile_payload["profiles"][st.session_state.active_user_profile]
        _prefs_by_sheet = _profile.setdefault("preferences_by_sheet", {})
        if "preferences" in _profile and "Analysis" not in _prefs_by_sheet:
            _prefs_by_sheet["Analysis"] = _profile.pop("preferences")
        _prefs_by_sheet[_selected_profile_sheet] = get_profile_preferences(_selected_profile_sheet)
        save_user_profiles(profile_payload)
    except Exception as profile_save_exc:
        st.warning(("Préférences non enregistrées dans Nextcloud : " if st.session_state.language == "fr" else "Preferences were not saved to Nextcloud: ") + str(profile_save_exc))
