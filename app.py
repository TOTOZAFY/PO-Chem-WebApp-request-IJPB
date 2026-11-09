import os
import re
import ssl
import smtplib
import json
import uuid
import unicodedata
from datetime import date, datetime
from email.message import EmailMessage
from io import BytesIO
from urllib.parse import quote

import requests
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
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
    "Métabolomique non ciblée (métabolites primaires) par GC-MS",
    "Métabolomique non ciblée (métabolites secondaires) par LC-HRMS",
    "Analyse des acides gras par GC-MS",
    "Analyse des stérols totaux par GC-MS",
    "Lipidomique non ciblée par LC-HRMS",
    "Lipidomique ciblée par LC-MS/MS",
    "Oligoglycomique par LC-HRMS",
    "Analyse des monosaccharides par GC-MS",
    "Analyse des hormones par LC-MS/MS",
    "Développement de méthodes spécifiques par GC-MS",
    "Développement de méthodes spécifiques par LC-MS/MS",
    "Développement de méthodes spécifiques par LC-HRMS",
]
SPECIFIC_DEVELOPMENT_TYPE = "Développement de méthodes spécifiques par LC-MS/MS"
ANALYSIS_TYPES_EN = {
    "Métabolomique non ciblée (métabolites primaires) par GC-MS": "Untargeted metabolomics (primary metabolites) by GC-MS",
    "Métabolomique non ciblée (métabolites secondaires) par LC-HRMS": "Untargeted metabolomics (secondary metabolites) by LC-HRMS",
    "Analyse des acides gras par GC-MS": "Fatty acid analysis by GC-MS",
    "Analyse des stérols totaux par GC-MS": "Total sterol analysis by GC-MS",
    "Lipidomique non ciblée par LC-HRMS": "Untargeted lipidomics by LC-HRMS",
    "Lipidomique ciblée par LC-MS/MS": "Targeted lipidomics by LC-MS/MS",
    "Oligoglycomique par LC-HRMS": "Oligoglycomics by LC-HRMS",
    "Analyse des monosaccharides par GC-MS": "Monosaccharide analysis by GC-MS",
    "Analyse des hormones par LC-MS/MS": "Hormone analysis by LC-MS/MS",
    "Développement de méthodes spécifiques par GC-MS": "Custom method development by GC-MS",
    "Développement de méthodes spécifiques par LC-MS/MS": "Custom method development by LC-MS/MS",
    "Développement de méthodes spécifiques par LC-HRMS": "Custom method development by LC-HRMS",
}

TRANSLATIONS = {
    "fr": {
        "language": "Langue / Language", "form": "Formulaire de demande",
        "admin": "Espace administrateur", "intro": "Complétez les rubriques ci-dessous. "
        "À la dernière étape, votre demande sera enregistrée dans Nextcloud de la plateforme et l'administrateur sera notifié. "
        "La feuille de demande pourra également être téléchargée si nécessaire après la soumission.",
        "steps": ["Contacts", "Description du projet", "Échantillons", "Résultats & bibliographie",
                  "Informations complémentaires", "Vérification & soumission"],
        "general": "1. Informations générales", "requester_section": "Demandeur", "administrator_section": "Plateforme Observatoire du végétal chimie et métabolisme", "manager": "Responsable du projet *", "manager_last_name": "Nom *", "manager_first_name": "Prénom *",
        "visibility": "Type de laboratoire *",
        "visibility_private": "Privée",
        "visibility_public": "Publique",
        "visibility_ijpb": "Publique (IJPB)",
        "group": "Groupe / laboratoire *", "date": "Date", "phone": "Téléphone",
        "manager_email": "E-mail du responsable *", "admin_staff": "Personnel administratif",
        "admin_phone": "Téléphone du personnel administratif", "admin_email": "E-mail du personnel administratif (facultatif)",
        "recommended": "* Champ obligatoire.", "description": "2. Description du projet",
        "description_help": "Questions scientifiques et techniques à traiter, molécules à analyser.",
        "description_field": "Description scientifique et technique",
        "samples_title": "3. Nature et nombre des échantillons",
        "samples_help": "Indiquez notamment les références du procédé analytique, les solvants et les formules chimiques si disponibles.",
        "samples_field": "Nature et informations analytiques",
        "sample_count": "Nombre d’échantillons *",
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
        "ijpb_admin_email": "E-mail du personnel administratif de l’IJPB (CC, facultatif)",
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
        "analysis_mode": "Mode d’ionisation / d’injection",
        "analysis_mode_lchrms": "Mode d’ionisation (ESI)",
        "analysis_mode_gc_ms": "Mode d’injection",
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
        "At the final step, your request will be saved on the platform's Nextcloud and the administrator will be notified. "
        "The request form can also be downloaded after submission if needed.",
        "steps": ["Contacts", "Project description", "Samples", "Preliminary results & bibliography",
                  "Additional information", "Review & submission"],
        "general": "1. General information", "requester_section": "Requester", "administrator_section": "Plant observatory chemistry and metabolism platform", "manager": "Project manager *", "manager_last_name": "Last name *", "manager_first_name": "First name *",
        "visibility": "Laboratory type *",
        "visibility_private": "Private",
        "visibility_public": "Public",
        "visibility_ijpb": "Public (IJPB)",
        "group": "Group / laboratory *", "date": "Date", "phone": "Telephone",
        "manager_email": "Project manager email *", "admin_staff": "Administrative staff",
        "admin_phone": "Administrative staff telephone", "admin_email": "Administrative staff email (optional)",
        "recommended": "* Required field.", "description": "2. Project description",
        "description_help": "Scientific and technical questions to be addressed, molecules to be analysed.",
        "description_field": "Scientific and technical description",
        "samples_title": "3. Nature and number of samples",
        "samples_help": "Please include analytical process references, solvents and chemical formulas if available.",
        "samples_field": "Nature and analytical information",
        "sample_count": "Number of samples *",
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
        "ijpb_admin_email": "IJPB administrative staff email (CC, optional)",
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
        "analysis_mode": "Ionization / injection mode",
        "analysis_mode_lchrms": "Ionization mode (ESI)",
        "analysis_mode_gc_ms": "Injection mode",
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


def normalize_analysis_type_name(value):
    """Correct known French labels and ensure every analysis type ends with an MS platform."""
    name = str(value or "").strip()
    corrections = {
        "Métabolites primaires par GC-MS": "Métabolomique non ciblée (métabolites primaires) par GC-MS",
        "Métabolomique non ciblée par LC-HRMS": "Métabolomique non ciblée (métabolites secondaires) par LC-HRMS",
        "Analyses d'acides gras par GC-MS": "Analyse des acides gras par GC-MS",
        "Analyse de stérols totaux par GC-MS": "Analyse des stérols totaux par GC-MS",
        "Lipidomique non ciblée LC-HRMS": "Lipidomique non ciblée par LC-HRMS",
        "Lipidomique ciblée LC-MS/MS": "Lipidomique ciblée par LC-MS/MS",
        "Oligoglycomique par LC-HRMS": "Oligoglycomique par LC-HRMS",
        "Monosaccharides par GC-MS": "Analyse des monosaccharides par GC-MS",
        "Hormones par LC-MS/MS": "Analyse des hormones par LC-MS/MS",
        "Développements spécifiques par GC-MS": "Développement de méthodes spécifiques par GC-MS",
        "Développements spécifiques par LC-MS/MS": "Développement de méthodes spécifiques par LC-MS/MS",
        "Développements spécifiques par LC-HRMS": "Développement de méthodes spécifiques par LC-HRMS",
        "Développements spécifiques": "Développement de méthodes spécifiques par LC-MS/MS",
    }
    name = corrections.get(name, name)
    if not name.endswith(("LC-HRMS", "GC-MS", "LC-MS/MS")):
        name += " par LC-MS/MS"
    return name


def load_notification_settings():
    """Load editable notification recipients from Nextcloud."""
    response = webdav_request("GET", NOTIFICATION_SETTINGS_PATH)
    if response.status_code == 404:
        return {
            "admin_email": DEFAULT_ADMIN_EMAIL,
            "admin_cc_email": "",
            "ijpb_admin_email": "stephanie.zimmermann@inrae.fr",
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
        "ijpb_admin_email": str(settings.get("ijpb_admin_email") or "stephanie.zimmermann@inrae.fr").strip(),
        "analysis_types": list(dict.fromkeys(
            ANALYSIS_TYPES + [
                normalize_analysis_type_name(x)
                for x in (
                    settings.get("analysis_types", ANALYSIS_TYPES)
                    if isinstance(settings.get("analysis_types", ANALYSIS_TYPES), list)
                    else ANALYSIS_TYPES
                )
                if str(x).strip() and normalize_analysis_type_name(x) not in ANALYSIS_TYPES
            ]
        )),
        "analysis_type_translations": {
            **(
                {
                    normalize_analysis_type_name(str(k).strip()): str(v).strip()
                    for k, v in settings.get("analysis_type_translations", {}).items()
                    if str(k).strip() and str(v).strip()
                }
                if isinstance(settings.get("analysis_type_translations", {}), dict)
                else {}
            ),
            # Keep the reviewed English labels authoritative for the 12 predefined types,
            # including when older French labels are still present in the Nextcloud JSON.
            **ANALYSIS_TYPES_EN,
        },
        "analysis_catalog": (
            settings.get("analysis_catalog", [])
            if isinstance(settings.get("analysis_catalog", []), list) else []
        ),
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

def filename_part(value, fallback):
    """Create a readable, filesystem-safe filename component (ASCII, hyphenated)."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^A-Za-z0-9]+", "-", value).strip("-")
    return value[:70] or fallback


def request_display_label(record, include_date=True):
    """Consistent requester-project-lab label for administrator request lists."""
    parts = [
        filename_part(record.get("project_manager"), "DEMANDEUR"),
        filename_part(record.get("project_name"), "PROJET"),
        filename_part(record.get("group"), "LABO"),
    ]
    request_id = str(record.get("id", "")).strip()
    label = " — ".join(([request_id] if request_id else []) + parts)
    if include_date:
        label += f" — {str(record.get('submitted_at', ''))[:10]}"
    return label


def build_document(form_data):
    doc = Document(TEMPLATE_PATH)
    if len(doc.tables) < 2:
        raise ValueError("Le modèle Word ne contient pas le tableau de formulaire attendu.")
    table = doc.tables[1]
    manager_cell = table.cell(0, 0)
    manager_cell.paragraphs[0].clear()
    manager_last = str(form_data.get("project_manager", "") or "").strip()
    manager_first = str(form_data.get("project_manager_first_name", "") or "").strip()
    manager_label = "Nom" if st.session_state.get("language", "fr") == "fr" else "Last name"
    first_label = "Prénom" if st.session_state.get("language", "fr") == "fr" else "First name"
    manager_p = manager_cell.paragraphs[0]
    manager_p.add_run(f"{manager_label} : ").bold = True
    manager_p.add_run(manager_last)
    first_p = manager_cell.add_paragraph()
    first_p.add_run(f"{first_label} : ").bold = True
    first_p.add_run(manager_first)
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
    sample_count = form_data.get("sample_count", "")
    if not form_data.get("analysis_items") and sample_count not in ("", None):
        count_paragraph = samples_cell.add_paragraph()
        count_paragraph.paragraph_format.space_after = Pt(4)
        count_run = count_paragraph.add_run(
            ("Nombre d’échantillons : " if st.session_state.get("language", "fr") == "fr" else "Number of samples: ")
            + str(sample_count)
        )
        count_run.bold = True
        count_run.font.size = Pt(9)
    analysis_items = form_data.get("analysis_items") or [{
        "analysis_type": form_data.get("analysis_type", ""),
        "analysis_mode": form_data.get("analysis_mode", ""),
        "sample_count": form_data.get("sample_count", 1),
        "analysis_details": form_data.get("analysis_details", ""),
    }]
    for number, analysis in enumerate(analysis_items, 1):
        analysis_type = str(analysis.get("analysis_type", "")).strip()
        analysis_mode = str(analysis.get("analysis_mode", "")).strip()
        analysis_details = str(analysis.get("analysis_details", "")).strip()
        sample_count = analysis.get("sample_count", 1)
        paragraph = samples_cell.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(4)
        label = (f"Analyse {number} — " if st.session_state.get("language", "fr") == "fr" else f"Analysis {number} — ")
        run = paragraph.add_run(label + analysis_type + f" | {sample_count} " + ("échantillon(s)" if st.session_state.get("language", "fr") == "fr" else "sample(s)"))
        run.bold = True
        run.font.size = Pt(9)
        if analysis_mode:
            mode_p = samples_cell.add_paragraph(("Mode : " if st.session_state.get("language", "fr") == "fr" else "Mode: ") + analysis_mode)
            mode_p.paragraph_format.left_indent = Inches(0.2)
        if analysis_details:
            detail_p = samples_cell.add_paragraph(("Détails : " if st.session_state.get("language", "fr") == "fr" else "Details: ") + analysis_details)
            detail_p.paragraph_format.left_indent = Inches(0.2)

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
    visibility = form_data.get("visibility", "Public (IJPB)").strip()
    if visibility:
        visibility_paragraph = project_cell.add_paragraph()
        visibility_run = visibility_paragraph.add_run(
            ("Type de laboratoire : " if st.session_state.get("language", "fr") == "fr" else "Laboratory type: ")
            + visibility
        )
        visibility_run.bold = True
        visibility_run.font.size = Pt(9)
    add_to_existing_paragraph(project_cell.paragraphs[0], form_data.get("project_name"), bold=True)
    add_to_existing_paragraph(project_cell.paragraphs[1], f"{t('answer_prefix')} {form_data.get('funding', 'NO')}", bold=True)
    guidance_cell = table.cell(7, 0)
    add_to_existing_paragraph(guidance_cell.paragraphs[0], f"{t('answer_prefix')} {form_data.get('guidance', 'NO')}", bold=True)

    # New naming convention: REQ-<unique identifier>-<date>-<time>-<requester>-<project>-<lab>.
    request_id = form_data.get("request_id", "REQ-SANS-ID")
    manager_name = filename_part(form_data.get("project_manager"), "DEMANDEUR")
    project_name = filename_part(form_data.get("project_name"), "PROJET")
    lab_name = filename_part(form_data.get("group"), "LABO")
    # request_id already contains the unique identifier and submission date/time.
    reference = f"{request_id}-{manager_name}-{project_name}-{lab_name}"
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


def remove_formatting_from_empty_cells(workbook):
    """Keep empty spreadsheet cells visually blank: no fill, borders or other styling."""
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows(
            min_row=1, max_row=worksheet.max_row,
            min_col=1, max_col=worksheet.max_column
        ):
            for cell in row:
                if cell.value is None:
                    cell._style = None


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
    ws["B4"] = " ; ".join(x.get("analysis_type", "") for x in record.get("analysis_items", [])) or record.get("analysis_type", "")
    remove_formatting_from_empty_cells(wb)
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
    remove_formatting_from_empty_cells(workbook)
    output = BytesIO()
    workbook.save(output)
    path = record.get("fair_path") or fair_remote_path(record)
    response = webdav_request("PUT", path, data=output.getvalue())
    response.raise_for_status()
    record["fair_path"] = path
    record["fair_updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")



def fair_sheet_completion_status(record):
    """Return per-sheet indicators for user-entered changes vs the FAIR template."""
    if not os.path.isfile(FAIR_TEMPLATE_PATH):
        raise FileNotFoundError("FAIR_template.xlsx is missing")
    template_wb = load_workbook(FAIR_TEMPLATE_PATH, data_only=False)
    fair_wb = load_fair_workbook(record)
    statuses = {}
    for sheet_name in ("Sample", "Analysis", "Data treatment", "Pipeline Step by Step"):
        if sheet_name not in fair_wb.sheetnames or sheet_name not in template_wb.sheetnames:
            statuses[sheet_name] = False
            continue
        ws = fair_wb[sheet_name]
        template_ws = template_wb[sheet_name]
        changed = False
        max_row = max(ws.max_row, template_ws.max_row)
        max_col = max(ws.max_column, template_ws.max_column)
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                if sheet_name == "Sample" and (row, col) in {(1, 2), (2, 2), (3, 2), (4, 2)}:
                    continue
                current = ws.cell(row, col).value
                original = template_ws.cell(row, col).value
                if current != original and current not in (None, ""):
                    changed = True
                    break
            if changed:
                break
        statuses[sheet_name] = changed
    return statuses


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
        warning = ("Le modèle de métadonnées est fourni à titre d'exemple. Veuillez adapter les champs à votre expérience. Vous pouvez ajouter des lignes pour toute information expérimentale supplémentaire et laisser vides les rubriques non applicables. Conservez les intitulés standards utilisés pour la correspondance automatique ; ne remplissez pas les lignes File ID Q-Tof ESI+ et File ID Q-Tof ESI-." if fr else "The provided metadata template is an example. Please adapt the metadata fields to your own experiment: keep the fields that are relevant, add rows for any additional experimental information, and leave non-applicable fields blank. Do not change the standard field names where they are used for automatic matching ; let the rows File ID Q-Tof ESI+ and File ID Q-Tof ESI- empty.")
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
    render_keep_preferences_button("Analysis")


def render_fair_data_treatment(record, records):
    """Edit the Data treatment worksheet through a guided form or a full POS/NEG table."""
    import pandas as pd

    fr = st.session_state.get("language", "fr") == "fr"
    st.title("FAIR – Data treatment")
    st.caption(("Demande : " if fr else "Request: ") + str(record.get("id", "")))
    st.info(
        "Renseignez les paramètres réellement utilisés. Les valeurs déjà présentes dans le modèle sont des exemples modifiables. Les champs non applicables peuvent rester vides."
        if fr else
        "Enter the parameters actually used. Values already present in the template are editable examples. Non-applicable fields may be left blank."
    )
    try:
        wb = load_fair_workbook(record)
        if "Data treatment" not in wb.sheetnames:
            st.error("La feuille Data treatment est absente du classeur FAIR." if fr else "The Data treatment sheet is missing from the FAIR workbook.")
            return
        ws = wb["Data treatment"]

        # The selected treatment workflow and free-text method are stored in
        # dedicated FAIR cells below the existing MZmine parameter block.
        method_cell = "B82"
        custom_text_cell = "B83"
        method_label = "Type de traitement des données" if fr else "Data treatment method"
        custom_label = "Description du traitement personnalisé" if fr else "Custom data treatment description"
        ws["A82"] = method_label
        ws["A83"] = custom_label
        saved_method = str(ws[method_cell].value or "MZmine")
        saved_custom_text = str(ws[custom_text_cell].value or "")
        method_options = ["Traitement de données par MZmine", "Traitement de données personnalisé"] if fr else ["MZmine data processing", "Custom data processing"]
        default_method = method_options[1] if saved_method.lower() in ("custom", "personnalisé", method_options[1].lower()) else method_options[0]
        selected_method = st.radio(
            "Choisir le type de traitement" if fr else "Choose the data treatment type",
            method_options,
            index=method_options.index(default_method),
            key=f"fair_data_treatment_pref_method_{record.get('id')}"
        )
        is_custom = selected_method == method_options[1]
        if is_custom:
            st.markdown("**" + custom_label + "**")
            custom_description = st.text_area(
                "Décrivez toutes les étapes, logiciels, versions, paramètres, normalisations et filtres utilisés. Vous pouvez coller votre protocole complet ici." if fr else
                "Describe all steps, software, versions, parameters, normalization and filters used. You can paste your complete protocol here.",
                value=saved_custom_text,
                height=320,
                key=f"fair_data_treatment_pref_custom_text_{record.get('id')}"
            )
            if st.button("Enregistrer le traitement personnalisé" if fr else "Save custom data treatment", type="primary", key=f"dt_custom_save_{record.get('id')}"):
                ws[method_cell] = "Personnalisé" if fr else "Custom"
                ws[custom_text_cell] = custom_description.strip()
                save_fair_workbook(record, wb)
                save_metadata(records)
                st.success("Le traitement personnalisé a été enregistré dans le classeur FAIR." if fr else "Custom data treatment has been saved to the FAIR workbook.")
            return
        # MZmine mode: keep both the guided web form and the full POS/NEG table.
        # Switching to custom never deletes the MZmine values already stored below.

        # General fields stored once, before the POS/NEG parameter blocks.
        general = {"B2": str(ws["B2"].value or ""), "B3": str(ws["B3"].value or "")}
        general_labels = {"B2": str(ws["A2"].value or "Data processing by"), "B3": str(ws["A3"].value or "Where is the data stored?")}

        # Each parameter keeps its exact Excel row and its POS/NEG value cells.
        entries = []
        section = "General"
        for row in range(4, ws.max_row + 1):
            left_section = ws.cell(row, 1).value
            right_section = ws.cell(row, 5).value
            if left_section and str(left_section).strip().lower() not in ("mode",):
                section = str(left_section).strip()
            label_pos = ws.cell(row, 2).value
            label_neg = ws.cell(row, 6).value
            pos_cell = ws.cell(row, 3)
            neg_cell = ws.cell(row, 7)
            # Some rows are section/mode headings merged across the value columns.
            # MergedCell objects are read-only; expose only actual writable value cells.
            pos_writable = not isinstance(pos_cell, MergedCell)
            neg_writable = not isinstance(neg_cell, MergedCell)
            if (label_pos is None and label_neg is None) or not (pos_writable or neg_writable):
                continue
            label = str(label_pos or label_neg).strip()
            entries.append({
                "row": row,
                "section": section,
                "parameter": label,
                "pos": "" if not pos_writable or pos_cell.value is None else str(pos_cell.value),
                "neg": "" if not neg_writable or neg_cell.value is None else str(neg_cell.value),
                "pos_writable": pos_writable,
                "neg_writable": neg_writable,
            })

        form_tab, table_tab = st.tabs(["Formulaire web" if fr else "Web form", "Tableau complet POS / NEG" if fr else "Full POS / NEG table"])
        with form_tab:
            st.subheader("Informations générales" if fr else "General information")
            general_values = {}
            for cell, label in general_labels.items():
                general_values[cell] = st.text_input(label, value=general[cell], key=f"fair_data_treatment_pref_general_{cell}")
            current_section = None
            form_values = {}
            for item in entries:
                if item["section"] != current_section:
                    current_section = item["section"]
                    st.markdown(f"**{current_section}**")
                with st.expander(item["parameter"], expanded=False):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.text_input("POS", value=item["pos"], key=f"fair_data_treatment_pref_{item['row']}_pos")
                        form_values[(item["row"], "pos")] = st.session_state.get(f"fair_data_treatment_pref_{item['row']}_pos", item["pos"])
                    with c2:
                        st.text_input("NEG", value=item["neg"], key=f"fair_data_treatment_pref_{item['row']}_neg")
                        form_values[(item["row"], "neg")] = st.session_state.get(f"fair_data_treatment_pref_{item['row']}_neg", item["neg"])
            if st.button("Enregistrer le formulaire Data Treatment" if fr else "Save Data Treatment form", type="primary", key=f"dt_form_save_{record.get('id')}"):
                ws[method_cell] = "MZmine"
                ws[custom_text_cell] = saved_custom_text
                ws["B2"] = general_values["B2"]
                ws["B3"] = general_values["B3"]
                for (row, mode), value in form_values.items():
                    target = ws.cell(row, 3 if mode == "pos" else 7)
                    if not isinstance(target, MergedCell):
                        target.value = value.strip() if isinstance(value, str) else value
                save_fair_workbook(record, wb)
                save_metadata(records)
                st.success("Data treatment enregistré dans le classeur FAIR." if fr else "Data treatment saved to the FAIR workbook.")

        with table_tab:
            st.subheader("Tableau complet des paramètres" if fr else "Complete parameter table")
            table_general = {}
            g1, g2 = st.columns(2)
            with g1:
                table_general["B2"] = st.text_input(general_labels["B2"], value=general["B2"], key="fair_data_treatment_pref_table_general_B2")
            with g2:
                table_general["B3"] = st.text_input(general_labels["B3"], value=general["B3"], key="fair_data_treatment_pref_table_general_B3")
            table_data = [{"Ligne Excel": x["row"], "Section": x["section"], "Paramètre": x["parameter"], "POS": x["pos"], "NEG": x["neg"]} for x in entries]
            edited = st.data_editor(
                pd.DataFrame(table_data),
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["Ligne Excel", "Section", "Paramètre"],
                column_config={
                    "Ligne Excel": st.column_config.NumberColumn("Ligne Excel", width="small"),
                    "Section": st.column_config.TextColumn("Section", width="medium"),
                    "Paramètre": st.column_config.TextColumn("Paramètre", width="large"),
                    "POS": st.column_config.TextColumn("POS", width="medium"),
                    "NEG": st.column_config.TextColumn("NEG", width="medium"),
                },
                key="fair_data_treatment_pref_full_table",
            )
            if st.button("Enregistrer le tableau Data Treatment" if fr else "Save Data Treatment table", type="primary", key=f"dt_table_save_{record.get('id')}"):
                ws[method_cell] = "MZmine"
                ws[custom_text_cell] = saved_custom_text
                ws["B2"] = table_general["B2"]
                ws["B3"] = table_general["B3"]
                for _, line in edited.iterrows():
                    row = int(line["Ligne Excel"])
                    pos_target = ws.cell(row, 3)
                    neg_target = ws.cell(row, 7)
                    if not isinstance(pos_target, MergedCell):
                        pos_target.value = "" if pd.isna(line["POS"]) else str(line["POS"])
                    if not isinstance(neg_target, MergedCell):
                        neg_target.value = "" if pd.isna(line["NEG"]) else str(line["NEG"])
                save_fair_workbook(record, wb)
                save_metadata(records)
                st.success("Data treatment enregistré dans le classeur FAIR." if fr else "Data treatment saved to the FAIR workbook.")
    except Exception as exc:
        st.error(("Impossible de charger ou d'enregistrer Data treatment : " if fr else "Could not load or save Data treatment: ") + str(exc))
    render_keep_preferences_button("Data treatment")


def render_fair_pipeline_step(record, records):
    """Choose an interactive pipeline table or explicitly leave the FAIR sheet empty."""
    import pandas as pd
    fr = st.session_state.get("language", "fr") == "fr"
    st.title("FAIR – Pipeline Step by Step")
    st.caption(("Demande : " if fr else "Request: ") + str(record.get("id", "")))
    try:
        wb = load_fair_workbook(record)
        if "Pipeline Step by Step" not in wb.sheetnames:
            st.error("La feuille Pipeline Step by Step est absente du classeur FAIR." if fr else "Pipeline Step by Step sheet is missing from the FAIR workbook.")
            return
        ws = wb["Pipeline Step by Step"]
        mode_cell, mode_value_cell = "D1", "E1"
        saved_mode = str(ws[mode_value_cell].value or "Tableau interactif")
        options = ["Tableau interactif", "Aucun (laisser la feuille vide)"] if fr else ["Interactive table", "None (leave the sheet empty)"]
        default = options[1] if saved_mode.casefold() in ("none", "aucun", options[1].casefold()) else options[0]
        selected = st.radio("Mode de renseignement" if fr else "Entry mode", options, index=options.index(default), key=f"fair_pipeline_step_pref_mode_{record.get('id')}")
        if selected == options[1]:
            st.info("La feuille Pipeline Step by Step sera entièrement vidée lors de l'enregistrement." if fr else "The Pipeline Step by Step worksheet will be completely cleared when saved.")
            if st.button("Enregistrer : feuille vide" if fr else "Save: empty sheet", type="primary", key=f"pipeline_none_save_{record.get('id')}"):
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.__class__.__name__ != "MergedCell":
                            cell.value = None
                ws[mode_cell] = "Mode"
                ws[mode_value_cell] = "None"
                save_fair_workbook(record, wb)
                save_metadata(records)
                st.success("La feuille Pipeline Step by Step est maintenant vide." if fr else "Pipeline Step by Step sheet is now empty.")
            return

        # If the user previously selected None, restore the original step labels
        # from the pristine template when switching back to the interactive table.
        if not any(ws.cell(r, 1).value is not None or ws.cell(r, 2).value is not None for r in range(1, ws.max_row + 1)):
            template_wb = load_workbook(FAIR_TEMPLATE_PATH)
            template_ws = template_wb["Pipeline Step by Step"]
            for r in range(1, template_ws.max_row + 1):
                for c in (1, 2):
                    ws.cell(r, c).value = template_ws.cell(r, c).value
        # Read the predefined workflow from the original FAIR template (steps in A, polarity in B).
        rows = []
        for r in range(1, ws.max_row + 1):
            step = ws.cell(r, 1).value
            polarity = ws.cell(r, 2).value
            if step is None and polarity is None:
                continue
            if r == 1 and str(step or "").strip().casefold() in ("step", "étape"):
                continue
            rows.append({
                "Ligne Excel": r,
                "Étape": str(step or ""),
                "Mode": str(polarity or ""),
                "Réalisé": str(ws.cell(r, 3).value or ""),
                "Détails / commentaires": str(ws.cell(r, 4).value or ""),
            })
        st.caption("Modifiez le statut et ajoutez des précisions pour chaque étape. Les étapes et modes POS/NEG du modèle sont conservés." if fr else "Edit the status and add details for each step. The template steps and POS/NEG modes are retained.")
        df = pd.DataFrame(rows)

        # Bulk status assignment: all workflow rows or a user-selected subset.
        st.subheader("Affectation groupée des états" if fr else "Bulk status assignment")
        bulk_scope = st.radio(
            "Appliquer l'état à" if fr else "Apply status to",
            ["Toutes les lignes", "Uniquement les étapes sélectionnées"] if fr else ["All rows", "Selected steps only"],
            horizontal=True, key=f"pipeline_bulk_scope_{record.get('id')}"
        )
        step_choices = {
            f"L{int(row['Ligne Excel'])} – {row['Étape']} ({row['Mode']})": int(row['Ligne Excel'])
            for row in rows
        }
        selected_labels = []
        if bulk_scope != ("Toutes les lignes" if fr else "All rows"):
            selected_labels = st.multiselect(
                "Étapes concernées" if fr else "Steps to update",
                list(step_choices), key=f"pipeline_bulk_rows_{record.get('id')}"
            )
        bulk_states = ["Oui", "Non", "Partiel", "N/A"] if fr else ["Yes", "No", "Partial", "N/A"]
        bulk_status = st.selectbox(
            "État à appliquer" if fr else "Status to apply", bulk_states,
            key=f"pipeline_bulk_status_{record.get('id')}"
        )
        if st.button("Appliquer et enregistrer l'état" if fr else "Apply and save status", type="secondary", key=f"pipeline_bulk_apply_{record.get('id')}"):
            if bulk_scope != ("Toutes les lignes" if fr else "All rows") and not selected_labels:
                st.warning("Sélectionnez au moins une étape." if fr else "Select at least one step.")
            else:
                target_rows = set(step_choices.values()) if bulk_scope == ("Toutes les lignes" if fr else "All rows") else {step_choices[label] for label in selected_labels}
                for row_number in target_rows:
                    ws.cell(row_number, 3).value = bulk_status
                ws[mode_cell] = "Mode"
                ws[mode_value_cell] = "Interactive table"
                save_fair_workbook(record, wb)
                save_metadata(records)
                st.success((f"État '{bulk_status}' appliqué à {len(target_rows)} ligne(s)." if fr else f"Status '{bulk_status}' applied to {len(target_rows)} row(s)."))
                st.rerun()

        edited = st.data_editor(
            df, use_container_width=True, hide_index=True, num_rows="fixed",
            disabled=["Ligne Excel", "Étape", "Mode"],
            column_config={
                "Ligne Excel": st.column_config.NumberColumn("Ligne", width="small"),
                "Étape": st.column_config.TextColumn("Étape" if fr else "Step", width="large"),
                "Mode": st.column_config.TextColumn("Mode", width="small"),
                "Réalisé": st.column_config.SelectboxColumn("État" if fr else "Status", options=["", "Oui", "Non", "Partiel", "N/A"] if fr else ["", "Yes", "No", "Partial", "N/A"], width="medium"),
                "Détails / commentaires": st.column_config.TextColumn("Détails / commentaires" if fr else "Details / comments", width="large"),
            }, key=f"fair_pipeline_step_pref_editor_{record.get('id')}"
        )
        if st.button("Enregistrer le tableau Pipeline Step by Step" if fr else "Save Pipeline Step by Step table", type="primary", key=f"pipeline_table_save_{record.get('id')}"):
            # Preserve original A/B template labels, update only editable columns and mode marker.
            for _, line in edited.iterrows():
                r = int(line["Ligne Excel"])
                ws.cell(r, 3).value = "" if pd.isna(line["Réalisé"]) else str(line["Réalisé"])
                ws.cell(r, 4).value = "" if pd.isna(line["Détails / commentaires"]) else str(line["Détails / commentaires"])
            ws[mode_cell] = "Mode"
            ws[mode_value_cell] = "Interactive table"
            save_fair_workbook(record, wb)
            save_metadata(records)
            st.success("Pipeline Step by Step enregistré dans le classeur FAIR." if fr else "Pipeline Step by Step saved to the FAIR workbook.")
    except Exception as exc:
        st.error(("Impossible de charger ou d'enregistrer Pipeline Step by Step : " if fr else "Could not load or save Pipeline Step by Step: ") + str(exc))
    render_keep_preferences_button("Pipeline Step by Step")


def render_fair_pipeline_access():
    fr = st.session_state.get("language", "fr") == "fr"
    records = load_metadata()
    active_records = [r for r in records if not r.get("deleted", False)]
    role = st.selectbox("Qui souhaite compléter Pipeline Step by Step ?" if fr else "Who is completing Pipeline Step by Step?", ["Demandeur", "Administrateur"] if fr else ["Requester", "Administrator"], key="fair_pipeline_role")
    if role in (["Demandeur"] if fr else ["Requester"]):
        with st.form("fair_pipeline_requester_access_form"):
            reference = st.text_input("Référence de la demande" if fr else "Request reference")
            email = st.text_input("Adresse e-mail du demandeur" if fr else "Requester email address")
            submitted = st.form_submit_button("Accéder à Pipeline Step by Step" if fr else "Open Pipeline Step by Step")
        if submitted:
            record = next((r for r in active_records if str(r.get("id", "")).strip().casefold() == reference.strip().casefold() and str(r.get("manager_email", "")).strip().casefold() == email.strip().casefold()), None)
            if not record:
                st.error("Référence ou adresse e-mail incorrecte." if fr else "Reference or email address is incorrect.")
                return
            st.session_state["fair_pipeline_requester_id"] = record.get("id")
            st.rerun()
        saved_id = st.session_state.get("fair_pipeline_requester_id")
        record = next((r for r in active_records if str(r.get("id")) == str(saved_id)), None) if saved_id else None
        if record:
            render_fair_pipeline_step(record, records)
        return
    password_expected = get_secret("ADMIN_PASSWORD", "")
    if not st.session_state.get("fair_pipeline_admin_authenticated", False):
        with st.form("fair_pipeline_admin_login_form"):
            password = st.text_input("Mot de passe administrateur" if fr else "Administrator password", type="password")
            login = st.form_submit_button("Se connecter" if fr else "Log in")
        if login and password == password_expected:
            st.session_state["fair_pipeline_admin_authenticated"] = True
            st.rerun()
        elif login:
            st.error("Mot de passe incorrect." if fr else "Incorrect password.")
        return
    if active_records:
        choices = {f"{r.get('id')} – {r.get('project_name', '')}": r for r in active_records}
        selected = st.selectbox("Sélectionner une demande" if fr else "Select a request", list(choices), key="fair_pipeline_admin_request")
        render_fair_pipeline_step(choices[selected], records)
    else:
        st.info("Aucune demande disponible." if fr else "No requests available.")


def render_fair_data_treatment_access():
    """Reuse the FAIR requester/admin access pattern for Data treatment."""
    fr = st.session_state.get("language", "fr") == "fr"
    records = load_metadata()
    active_records = [r for r in records if not r.get("deleted", False)]
    role = st.selectbox(
        "Qui souhaite compléter la feuille Data treatment ?" if fr else "Who is completing the Data treatment sheet?",
        options=["Demandeur", "Administrateur"] if fr else ["Requester", "Administrator"],
        key="fair_data_treatment_role",
    )
    if role == ("Demandeur" if fr else "Requester"):
        with st.form("fair_data_treatment_requester_access_form"):
            reference = st.text_input("Référence de la demande" if fr else "Request reference")
            email = st.text_input("Adresse e-mail du demandeur" if fr else "Requester email address")
            submitted = st.form_submit_button("Accéder à Data treatment" if fr else "Open Data treatment")
        if submitted:
            record = next((r for r in active_records if str(r.get("id", "")).strip().casefold() == reference.strip().casefold() and str(r.get("manager_email", "")).strip().casefold() == email.strip().casefold()), None)
            if not record:
                st.error("Référence ou adresse e-mail incorrecte." if fr else "Request reference or email address is incorrect.")
                return
            st.session_state["fair_data_treatment_requester_id"] = record.get("id")
            st.rerun()
        saved_id = st.session_state.get("fair_data_treatment_requester_id")
        record = next((r for r in active_records if str(r.get("id")) == str(saved_id)), None) if saved_id else None
        if record:
            render_fair_data_treatment(record, records)
        return
    admin_password = get_secret("ADMIN_PASSWORD", "")
    if not admin_password:
        st.error("Le secret ADMIN_PASSWORD n'est pas configuré." if fr else "ADMIN_PASSWORD secret is not configured.")
        return
    if not st.session_state.get("fair_data_treatment_admin_authenticated", False):
        with st.form("fair_data_treatment_admin_login_form"):
            password = st.text_input("Mot de passe administrateur" if fr else "Administrator password", type="password")
            login = st.form_submit_button("Se connecter" if fr else "Log in")
        if login and password == admin_password:
            st.session_state["fair_data_treatment_admin_authenticated"] = True
            st.rerun()
        elif login:
            st.error("Mot de passe incorrect." if fr else "Incorrect password.")
        return
    if active_records:
        choices = {f"{r.get('id')} – {r.get('project_name', '')}": r for r in active_records}
        selected = st.selectbox("Sélectionner une demande" if fr else "Select a request", list(choices), key="fair_data_treatment_admin_request")
        render_fair_data_treatment(choices[selected], records)
    else:
        st.info("Aucune demande disponible." if fr else "No requests available.")


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

    if selected_sheet == "Data treatment":
        st.divider()
        render_fair_profile_manager("Data treatment")
        st.divider()
        render_fair_data_treatment_access()
        return

    if selected_sheet == "Pipeline Step by Step":
        st.divider()
        render_fair_profile_manager(selected_sheet)
        st.divider()
        render_fair_pipeline_access()
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
        ijpb_admin_email = st.text_input(
            t("ijpb_admin_email"),
            value=notification_settings.get("ijpb_admin_email", "stephanie.zimmermann@inrae.fr"),
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
                    "ijpb_admin_email": ijpb_admin_email.strip(),
                    "analysis_types": notification_settings.get("analysis_types", ANALYSIS_TYPES),
                    "analysis_type_translations": notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN),
                    "analysis_catalog": notification_settings.get("analysis_catalog", []),
                })
                st.success(t("notification_settings_saved"))
                st.rerun()
            except Exception as exc:
                st.error(f"{t('cloud_error')} {exc}")

    st.subheader(t("manage_analysis_types"))
    is_fr = st.session_state.language == "fr"
    managed_types = notification_settings.get("analysis_types") or ANALYSIS_TYPES.copy()
    type_translations = notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN.copy())
    stored_catalog = notification_settings.get("analysis_catalog", [])
    catalog_by_type = {
        normalize_analysis_type_name(item.get("type", "")): item
        for item in stored_catalog
        if isinstance(item, dict) and str(item.get("type", "")).strip()
    }
    type_translations = {
        normalize_analysis_type_name(key): value
        for key, value in type_translations.items()
    }
    def numeric_price(value):
        """Return a JSON-safe numeric price; keep empty cells as None."""
        if value in (None, ""):
            return None
        if isinstance(value, (int, float)):
            return float(value)
        # Migrate legacy text prices such as "125,50 €" to numeric values.
        cleaned = str(value).strip().replace(" ", "").replace(",", ".")
        cleaned = re.sub(r"[^0-9.\\-]", "", cleaned)
        try:
            return float(cleaned) if cleaned else None
        except ValueError:
            return None

    type_column = "Type d'analyse" if is_fr else "Analysis type"
    public_column = "Prix public" if is_fr else "Public price"
    private_column = "Prix privé" if is_fr else "Private price"
    inrae_column = "Prix public (IJPB)" if is_fr else "Public price (IJPB)"
    catalog_rows = []
    for analysis_type in managed_types:
        saved = catalog_by_type.get(analysis_type, {})
        translated_type = type_translations.get(
            analysis_type, ANALYSIS_TYPES_EN.get(analysis_type, analysis_type)
        )
        catalog_rows.append({
            type_column: analysis_type if is_fr else translated_type,
            "_canonical_type": analysis_type,
            public_column: numeric_price(saved.get("public_price")),
            private_column: numeric_price(saved.get("private_price")),
            inrae_column: numeric_price(saved.get("inrae_public_price")),
        })

    st.caption(
        "Correspondance avec le type de laboratoire : Privée → Prix privé ; Publique → Prix public ; Publique (IJPB) → Prix public (IJPB). "
        "Saisissez les tarifs unitaires en euros. Les cellules vides correspondent à un tarif non renseigné."
        if is_fr else
        "Mapping to laboratory type: Private → Private price; Public → Public price; Public (IJPB) → Public price (IJPB). "
        "Enter unit prices in euros. Empty cells mean that no price has been set."
    )
    edited_catalog = st.data_editor(
        catalog_rows,
        key="analysis_tariff_catalog_editor",
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            type_column: st.column_config.TextColumn(type_column, required=True),
            "_canonical_type": None,
            public_column: st.column_config.NumberColumn(
                "Prix public (€)" if is_fr else "Public price (€)",
                min_value=0.0, step=0.01, format="%.2f €"
            ),
            private_column: st.column_config.NumberColumn(
                "Prix privé (€)" if is_fr else "Private price (€)",
                min_value=0.0, step=0.01, format="%.2f €"
            ),
            inrae_column: st.column_config.NumberColumn(
                "Prix public (IJPB) (€)" if is_fr else "Public price (IJPB) (€)",
                min_value=0.0, step=0.01, format="%.2f €"
            ),
        },
    )
    col_save, col_add, col_edit, col_delete = st.columns([1.2, 1, 1, 1])
    with col_save:
        save_catalog_clicked = st.button(
            "Enregistrer le tableau" if is_fr else "Save catalog",
            type="primary", key="save_analysis_tariff_catalog"
        )
    with col_add:
        add_catalog_clicked = st.button(
            "Ajouter un type" if is_fr else "Add analysis type",
            key="add_analysis_catalog_type"
        )
    with col_edit:
        edit_catalog_clicked = st.button(
            "Modifier un nom" if is_fr else "Rename an analysis type",
            key="edit_analysis_catalog_type"
        )
    with col_delete:
        delete_catalog_clicked = st.button(
            "Supprimer un type" if is_fr else "Delete analysis type",
            key="delete_analysis_catalog_type"
        )

    def edited_catalog_rows():
        """Return data_editor output as a list of row dictionaries."""
        if hasattr(edited_catalog, "to_dict"):
            return edited_catalog.to_dict(orient="records")
        return [dict(row) for row in edited_catalog]

    def persist_analysis_catalog(rows):
        cleaned_types = []
        cleaned_catalog = []
        cleaned_translations = dict(type_translations)
        for row in rows:
            display_label = str(row.get(type_column, "")).strip()
            canonical_label = display_label if is_fr else str(row.get("_canonical_type", "")).strip()
            if not canonical_label or not display_label:
                continue
            if canonical_label in cleaned_types:
                raise ValueError("Un type d'analyse est présent plusieurs fois." if is_fr else "An analysis type is duplicated.")
            cleaned_types.append(canonical_label)
            cleaned_catalog.append({
                "type": canonical_label,
                "public_price": numeric_price(row.get(public_column)),
                "private_price": numeric_price(row.get(private_column)),
                "inrae_public_price": numeric_price(row.get(inrae_column)),
            })
            if not is_fr:
                cleaned_translations[canonical_label] = display_label
            else:
                cleaned_translations.setdefault(canonical_label, ANALYSIS_TYPES_EN.get(canonical_label, canonical_label))
        save_notification_settings({
            "admin_email": notification_settings.get("admin_email", DEFAULT_ADMIN_EMAIL),
            "admin_cc_email": notification_settings.get("admin_cc_email", ""),
            "ijpb_admin_email": notification_settings.get("ijpb_admin_email", "stephanie.zimmermann@inrae.fr"),
            "analysis_types": cleaned_types,
            "analysis_type_translations": cleaned_translations,
            "analysis_catalog": cleaned_catalog,
        })

    if save_catalog_clicked:
        try:
            persist_analysis_catalog(edited_catalog_rows())
            st.success("Catalogue tarifaire enregistré." if is_fr else "Tariff catalog saved.")
            st.rerun()
        except Exception as exc:
            st.error(f"{t('cloud_error')} {exc}")

    if add_catalog_clicked:
        st.session_state["show_add_analysis_catalog_type"] = True
    if st.session_state.get("show_add_analysis_catalog_type", False):
        st.info(
            "Important : chaque type d'analyse doit se terminer par une méthode analytique : LC-HRMS, GC-MS ou LC-MS/MS. Choisissez le suffixe ci-dessous ; il sera ajouté automatiquement."
            if is_fr else
            "Important: every analysis type must end with an analytical method: LC-HRMS, GC-MS or LC-MS/MS. Select the suffix below; it will be added automatically."
        )
        with st.form("new_analysis_catalog_type_form"):
            new_type = st.text_input("Nom du type d'analyse (sans la méthode à la fin)" if is_fr else "Analysis type name (without the method suffix)")
            new_type_method = st.selectbox(
                "Méthode analytique à ajouter à la fin" if is_fr else "Analytical method suffix",
                options=["LC-HRMS", "GC-MS", "LC-MS/MS"],
            )
            new_type_en = st.text_input("Nom anglais (sans la méthode à la fin)" if is_fr else "English name (without the method suffix)")
            confirm_add_type = st.form_submit_button("Ajouter" if is_fr else "Add")
        if confirm_add_type:
            base_candidate = new_type.strip().rstrip(" -")
            candidate = normalize_analysis_type_name(f"{base_candidate} {new_type_method}") if base_candidate else ""
            if not candidate:
                st.warning("Le nom du type est obligatoire." if is_fr else "Analysis type name is required.")
            elif any(candidate.casefold() == item.casefold() for item in managed_types):
                st.warning(t("analysis_type_exists"))
            else:
                try:
                    updated_rows = edited_catalog_rows()
                    updated_rows.append({
                        type_column: candidate,
                        "_canonical_type": candidate,
                        public_column: None,
                        private_column: None,
                        inrae_column: None,
                    })
                    if new_type_en.strip():
                        english_base = new_type_en.strip().rstrip(" -")
                        type_translations[candidate] = f"{english_base} {new_type_method}"
                    persist_analysis_catalog(updated_rows)
                    st.session_state.pop("show_add_analysis_catalog_type", None)
                    st.success("Type d'analyse ajouté." if is_fr else "Analysis type added.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"{t('cloud_error')} {exc}")

    if edit_catalog_clicked:
        st.session_state["show_edit_analysis_catalog_type"] = True
    if st.session_state.get("show_edit_analysis_catalog_type", False) and managed_types:
        edit_labels = {
            item: item if is_fr else type_translations.get(item, ANALYSIS_TYPES_EN.get(item, item))
            for item in managed_types
        }
        with st.form("edit_analysis_catalog_type_form"):
            selected_edit_label = st.selectbox(
                "Type à modifier" if is_fr else "Analysis type to rename",
                options=list(edit_labels.values()),
            )
            type_to_edit = next((key for key, value in edit_labels.items() if value == selected_edit_label), selected_edit_label)
            current_english = type_translations.get(type_to_edit, ANALYSIS_TYPES_EN.get(type_to_edit, type_to_edit))
            renamed_fr = st.text_input("Nouveau nom français" if is_fr else "New French name", value=type_to_edit)
            renamed_en = st.text_input("Nouveau nom anglais" if is_fr else "New English name", value=current_english)
            confirm_rename_type = st.form_submit_button("Enregistrer le nouveau nom" if is_fr else "Save new name")
        if confirm_rename_type:
            new_canonical = normalize_analysis_type_name(renamed_fr)
            if not renamed_fr.strip():
                st.warning("Le nouveau nom est obligatoire." if is_fr else "The new name is required.")
            elif any(x != type_to_edit and x.casefold() == new_canonical.casefold() for x in managed_types):
                st.warning(t("analysis_type_exists"))
            else:
                try:
                    rows = edited_catalog_rows()
                    for row in rows:
                        row_canonical = str(row.get("_canonical_type", row.get(type_column, ""))).strip()
                        if row_canonical == type_to_edit:
                            row["_canonical_type"] = new_canonical
                            row[type_column] = new_canonical if is_fr else (renamed_en.strip() or new_canonical)
                    type_translations.pop(type_to_edit, None)
                    type_translations[new_canonical] = renamed_en.strip() or new_canonical
                    persist_analysis_catalog(rows)
                    st.session_state.pop("show_edit_analysis_catalog_type", None)
                    st.success("Nom du type d'analyse modifié." if is_fr else "Analysis type renamed.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"{t('cloud_error')} {exc}")

    removable_types = managed_types.copy()
    if delete_catalog_clicked and removable_types:
        st.session_state["show_delete_analysis_catalog_type"] = True
    if st.session_state.get("show_delete_analysis_catalog_type", False) and removable_types:
        with st.form("delete_analysis_catalog_type_form"):
            remove_labels = {
                item: item if is_fr else type_translations.get(item, ANALYSIS_TYPES_EN.get(item, item))
                for item in removable_types
            }
            selected_remove_label = st.selectbox(
                "Type à supprimer" if is_fr else "Type to delete",
                options=list(remove_labels.values()),
            )
            type_to_remove = next(
                (canonical for canonical, translated in remove_labels.items()
                 if translated == selected_remove_label),
                selected_remove_label,
            )
            confirm_delete_type = st.form_submit_button("Confirmer la suppression" if is_fr else "Confirm deletion")
        if confirm_delete_type:
            try:
                remaining_rows = [
                    row for row in edited_catalog_rows()
                    if str(row.get("_canonical_type", row.get(type_column, ""))).strip() != type_to_remove
                ]
                persist_analysis_catalog(remaining_rows)
                st.session_state.pop("show_delete_analysis_catalog_type", None)
                st.success("Type d'analyse supprimé." if is_fr else "Analysis type deleted.")
                st.rerun()
            except Exception as exc:
                st.error(f"{t('cloud_error')} {exc}")

    # Archived requests remain in the same Nextcloud metadata file; only their
    # archived flag changes. This preserves existing records and compatibility.
    active_records = [r for r in records if not r.get("archived", False)]
    archived_records = [r for r in records if r.get("archived", False)]

    st.subheader("Recettes annuelles indicatives" if st.session_state.language == "fr" else "Indicative annual revenue")
    st.warning(
        "⚠️ ESTIMATION INDICATIVE UNIQUEMENT — Les montants sont calculés à partir du catalogue tarifaire actuel. Ils ne représentent pas nécessairement les sommes réellement facturées ou encaissées."
        if st.session_state.language == "fr" else
        "⚠️ INDICATIVE ESTIMATE ONLY — Amounts are calculated using the current price catalog. They do not necessarily represent amounts actually invoiced or received."
    )
    st.caption(
        "Cumul de toutes les demandes acceptées, actives et archivées."
        if st.session_state.language == "fr" else
        "Total of all accepted requests, active and archived."
    )
    catalog_prices = {
        normalize_analysis_type_name(item.get("type", "")): item
        for item in notification_settings.get("analysis_catalog", [])
        if isinstance(item, dict) and str(item.get("type", "")).strip()
    }
    # Keep stored laboratory categories unchanged for compatibility with existing requests.
    # IJPB revenue is split according to the request's funding checkbox.
    revenue_categories = (
        "Privée",
        "Public",
        "Public (IJPB) - funded",
        "Public (IJPB) - not funded",
    )
    revenue_labels = {
        "Privée": "Privée" if is_fr else "Private",
        "Public": "Publique" if is_fr else "Public",
        "Public (IJPB) - funded": (
            "Publique (IJPB) — projet financé" if is_fr else "Public (IJPB) — funded project"
        ),
        "Public (IJPB) - not funded": (
            "Publique (IJPB) — projet non financé" if is_fr else "Public (IJPB) — non-funded project"
        ),
    }
    annual_revenue = {}
    annual_request_counts = {}
    missing_price_count = 0
    for item in records:
        if item.get("deleted", False) or item.get("status") != "accepted":
            continue
        decision_date = str(item.get("decision_at", "") or item.get("submitted_at", ""))
        year = decision_date[:4]
        if not year.isdigit():
            continue
        analysis_items = item.get("analysis_items") or [{
            "analysis_type": item.get("analysis_type", ""),
            "analysis_mode": item.get("analysis_mode", ""),
            "sample_count": item.get("sample_count", 1),
        }]
        lab_type = str(item.get("visibility", "Public (IJPB)") or "Public (IJPB)")
        lab_type = {"Publique": "Public", "Publique (IJPB)": "Public (IJPB)"}.get(lab_type, lab_type)
        price_key = {"Privée": "private_price", "Public": "public_price", "Public (IJPB)": "inrae_public_price"}.get(lab_type)
        revenue_category = lab_type
        if lab_type == "Public (IJPB)":
            funding_value = str(item.get("funding", "NO") or "NO").strip().upper()
            revenue_category = "Public (IJPB) - funded" if funding_value in ("YES", "OUI", "TRUE", "1") else "Public (IJPB) - not funded"
        year_totals = annual_revenue.setdefault(year, {category: 0.0 for category in revenue_categories})
        for analysis in analysis_items:
            price_record = catalog_prices.get(normalize_analysis_type_name(str(analysis.get("analysis_type", ""))), {})
            raw_price = price_record.get(price_key) if price_key else None
            try:
                unit_price = float(str(raw_price).replace(",", ".")) if raw_price not in (None, "") else None
                sample_count = max(1, int(analysis.get("sample_count", 1) or 1))
            except (TypeError, ValueError):
                unit_price, sample_count = None, 1
            multiplier = 2 if str(analysis.get("analysis_mode", "")) in ("Split/Splitless", "ESI+/-") else 1
            if unit_price is None:
                missing_price_count += 1
                continue
            year_totals[revenue_category] += unit_price * sample_count * multiplier
        annual_request_counts[year] = annual_request_counts.get(year, 0) + 1

    if annual_revenue:
        import pandas as pd
        year_col = "Année" if is_fr else "Year"
        total_col = "Total indicatif (€)" if is_fr else "Indicative total (€)"
        category_cols = {category: f"{revenue_labels[category]} (€)" for category in revenue_categories}
        revenue_frame = pd.DataFrame([
            {
                year_col: year,
                **{category_cols[category]: round(amounts[category], 2) for category in revenue_categories},
                total_col: round(sum(amounts.values()), 2),
            }
            for year, amounts in sorted(annual_revenue.items())
        ])
        st.bar_chart(revenue_frame.set_index(year_col)[list(category_cols.values())])
        st.dataframe(revenue_frame, use_container_width=True, hide_index=True)
        category_totals = {
            category: sum(amounts[category] for amounts in annual_revenue.values())
            for category in revenue_categories
        }
        category_metrics = st.columns(4)
        for column, category in zip(category_metrics, revenue_categories):
            column.metric(
                revenue_labels[category],
                f"{category_totals[category]:,.2f} €".replace(",", " "),
            )
        total_estimated = sum(category_totals.values())
        st.metric(
            "Total indicatif cumulé" if is_fr else "Cumulative indicative total",
            f"{total_estimated:,.2f} €".replace(",", " ")
        )
        count_label = "Demandes acceptées comptabilisées" if is_fr else "Accepted requests included"
        st.caption(f"{count_label} : {sum(annual_request_counts.values())}")
    else:
        st.info("Aucune demande acceptée avec un tarif renseigné et une année d'acceptation identifiable." if is_fr else "No accepted requests with a catalog price and identifiable acceptance year.")
    if missing_price_count:
        st.warning(
            f"{missing_price_count} demande(s) acceptée(s) exclue(s) du calcul car aucun tarif correspondant n'est renseigné."
            if is_fr else
            f"{missing_price_count} accepted request(s) excluded because no matching catalog price is configured."
        )

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
                  "request_date", "submitted_at", "status", "admin_comment", "document_path", "quote_signed", "signed_quote_path", "analysis_type", "analysis_mode", "analysis_details"]
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
                f"{request_display_label(r)} — {r.get('id', '')}": r["id"]
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

            # Operator assigned to perform the analyses for this request.
            operator_label = "Opérateur en charge de l’analyse" if st.session_state.language == "fr" else "Operator responsible for the analysis"
            operator_save_label = "Enregistrer l’opérateur" if st.session_state.language == "fr" else "Save assigned operator"
            operator_value = st.text_input(
                operator_label,
                value=record.get("assigned_operator", ""),
                key=f"assigned_operator_input_{selected_id}",
                placeholder="Nom de l’opérateur" if st.session_state.language == "fr" else "Operator name",
            )
            operator_email_label = (
                "Adresse e-mail de la personne en charge de l’analyse"
                if st.session_state.language == "fr"
                else "Email address of the person responsible for the analysis"
            )
            operator_email_value = st.text_input(
                operator_email_label,
                value=record.get("assigned_operator_email", ""),
                key=f"assigned_operator_email_input_{selected_id}",
                placeholder="nom@inrae.fr",
                help=(
                    "Cette adresse sera mise en copie de la notification envoyée au demandeur lorsque la demande sera acceptée."
                    if st.session_state.language == "fr"
                    else "This address will be copied on the notification sent to the requester when the request is accepted."
                ),
            )
            operator_email_is_valid = (
                not operator_email_value.strip()
                or bool(re.fullmatch(r"[^\\s@]+@[^\\s@]+\\.[^\\s@]+", operator_email_value.strip()))
            )
            if not operator_email_is_valid:
                st.warning(
                    "Veuillez saisir une adresse e-mail valide."
                    if st.session_state.language == "fr"
                    else "Please enter a valid email address."
                )
            if st.button(operator_save_label, key=f"save_assigned_operator_{selected_id}"):
                if not operator_email_is_valid:
                    st.error(
                        "L’adresse e-mail de la personne en charge est invalide."
                        if st.session_state.language == "fr"
                        else "The responsible person's email address is invalid."
                    )
                else:
                    record["assigned_operator"] = operator_value.strip()
                    record["assigned_operator_email"] = operator_email_value.strip()
                    record["operator_assigned_at"] = datetime.now().astimezone().isoformat(timespec="seconds") if operator_value.strip() else ""
                    try:
                        save_metadata(records)
                        st.success("Opérateur et adresse e-mail enregistrés pour cette demande." if st.session_state.language == "fr" else "Operator and email address saved for this request.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"{t('cloud_error')} {exc}")

            stored_type = record.get("analysis_type", "")
            shown_type = stored_type if st.session_state.language == "fr" else notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN).get(stored_type, ANALYSIS_TYPES_EN.get(stored_type, stored_type))
            admin_analysis_items = record.get("analysis_items") or [{"analysis_type": stored_type, "analysis_mode": record.get("analysis_mode", ""), "sample_count": record.get("sample_count", 1), "analysis_details": record.get("analysis_details", "")}]
            for analysis_number, analysis in enumerate(admin_analysis_items, 1):
                raw_type = analysis.get("analysis_type", "")
                display_type = raw_type if st.session_state.language == "fr" else notification_settings.get("analysis_type_translations", ANALYSIS_TYPES_EN).get(raw_type, ANALYSIS_TYPES_EN.get(raw_type, raw_type))
                st.write(f"**Analyse {analysis_number if st.session_state.language == 'fr' else 'Analysis ' + str(analysis_number)} :** {display_type} — {analysis.get('sample_count', 1)} échantillon(s)")
                if analysis.get("analysis_mode"):
                    st.caption(f"{t('analysis_mode')} : {analysis.get('analysis_mode')}")
                if analysis.get("analysis_details"):
                    st.caption(f"{t('specific_development_details')} : {analysis.get('analysis_details')}")
            st.write(f"**{t('doc_path')} :** `{record.get('document_path','')}`")
            request_status = record.get("status", "pending")
            request_status_labels = {
                "pending": t("status_pending"),
                "accepted": t("status_accepted"),
                "refused": t("status_refused"),
            }
            status_label = "Statut de la demande" if st.session_state.language == "fr" else "Request status"
            st.write(f"**{status_label} :** {request_status_labels.get(request_status, request_status or '—')}")
            operator_status_label = "Opérateur en charge de l’analyse" if st.session_state.language == "fr" else "Operator responsible for the analysis"
            operator_status_value = record.get("assigned_operator", "").strip() or ("Non renseigné" if st.session_state.language == "fr" else "Not specified")
            st.write(f"**{operator_status_label} :** {operator_status_value}")
            quotation_imported = bool(record.get("signed_quote_path"))
            quotation_label = "Devis importé" if st.session_state.language == "fr" else "Quotation uploaded"
            quotation_value = ("Oui" if st.session_state.language == "fr" else "Yes") if quotation_imported else ("Non" if st.session_state.language == "fr" else "No")
            st.write(f"**{quotation_label} :** {quotation_value}")

            fair_status_heading = "État des feuilles FAIR" if st.session_state.language == "fr" else "FAIR sheet completion"
            st.markdown(f"**{fair_status_heading}**")
            try:
                sheet_statuses = fair_sheet_completion_status(record)
                sheet_names_fr = {
                    "Sample": "Sample",
                    "Analysis": "Analysis",
                    "Data treatment": "Data treatment",
                    "Pipeline Step by Step": "Pipeline Step by Step",
                }
                for sheet_name, has_content in sheet_statuses.items():
                    sheet_label = sheet_names_fr[sheet_name]
                    if st.session_state.language == "fr":
                        state_text = "Données saisies" if has_content else "Non renseignée"
                    else:
                        state_text = "Data entered" if has_content else "Not completed"
                    st.write(f"- **{sheet_label} :** {state_text}")
            except Exception as fair_status_error:
                st.caption(("État des feuilles FAIR indisponible : " if st.session_state.language == "fr" else "FAIR sheet status unavailable: ") + str(fair_status_error))

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
                                    "Bonjour " + (record.get("project_manager_first_name") or record.get("project_manager", "")) + ",\n\n"
                                    "Merci de compléter les métadonnées FAIR de votre demande.\n"
                                    "Référence de votre demande : " + str(record.get("id", "")) + "\n"
                                    "Vous pouvez ouvrir l'application, choisir le menu FAIR, puis saisir cette référence et votre adresse e-mail de demandeur.\n\n"
                                    "Accès direct sécurisé :\n" + link
                                )
                            else:
                                body = (
                                    "Hello " + (record.get("project_manager_first_name") or record.get("project_manager", "")) + ",\n\n"
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
                        # On acceptance, enrich the requester notification with the
                        # request details and the applicable catalogue price.
                        recipients_settings = load_notification_settings()
                        cc_addresses = [record.get("admin_email", ""), record.get("platform_staff_email", "")]
                        if new_status == "accepted":
                            cc_addresses.append(recipients_settings.get("ijpb_admin_email", "stephanie.zimmermann@inrae.fr"))
                            cc_addresses.append(record.get("assigned_operator_email", ""))
                            visibility = str(record.get("visibility", "Public (IJPB)") or "Public (IJPB)")
                            analysis_items = record.get("analysis_items") or [{
                                "analysis_type": record.get("analysis_type", ""),
                                "analysis_mode": record.get("analysis_mode", ""),
                                "sample_count": record.get("sample_count", 1),
                            }]
                            catalog = {
                                normalize_analysis_type_name(str(x.get("type", ""))): x
                                for x in recipients_settings.get("analysis_catalog", []) if isinstance(x, dict)
                            }
                            price_key = {"Privée": "private_price", "Public": "public_price", "Public (IJPB)": "inrae_public_price"}.get(visibility)
                            total_price = 0.0
                            price_complete = True
                            details_fr, details_en = [], []
                            for analysis in analysis_items:
                                atype = str(analysis.get("analysis_type", ""))
                                mode = str(analysis.get("analysis_mode", "") or "")
                                count = max(1, int(analysis.get("sample_count", 1) or 1))
                                multiplier = 2 if mode in ("Split/Splitless", "ESI+/-") else 1
                                entry = catalog.get(normalize_analysis_type_name(atype), {})
                                raw = entry.get(price_key) if price_key else None
                                try:
                                    unit = float(str(raw).strip().replace(" ", "").replace(",", ".").replace("€", "")) if raw not in (None, "") else None
                                except (TypeError, ValueError):
                                    unit = None
                                if unit is None:
                                    price_complete = False
                                else:
                                    total_price += unit * count * multiplier
                                details_fr.append(f"- {atype} | {count} échantillon(s) | {mode or 'Mode non précisé'}")
                                details_en.append(f"- {atype} | {count} sample(s) | {mode or 'Mode not specified'}")
                            sample_count = sum(max(1, int(x.get("sample_count", 1) or 1)) for x in analysis_items)
                            analysis_details_fr = "\n".join(details_fr)
                            analysis_details_en = "\n".join(details_en)
                            price_text_fr = f"{total_price:.2f} €" if price_complete else "Indisponible : au moins un tarif manque dans le catalogue."
                            price_text_en = f"€{total_price:.2f}" if price_complete else "Unavailable: at least one catalog price is missing."
                            if visibility in ("Public", "Privée"):
                                quote_followup_fr = "Un devis détaillé vous sera envoyé prochainement."
                                quote_followup_en = "A detailed quotation will be sent to you shortly."
                            else:
                                quote_followup_fr = ""
                                quote_followup_en = ""

                            disclaimer_fr = "Le prix indiqué est donné à titre indicatif. Seul le devis signé fait foi."
                            disclaimer_en = "The displayed price is indicative only. Only the signed quotation is authoritative."

                            if st.session_state.language == "fr":
                                body = (
                                    f"Bonjour {record.get('project_manager_first_name') or record.get('project_manager', '')},\n\n"
                                    f"Votre demande d'analyse a été acceptée.\n\n"
                                    f"Référence de la demande : {record.get('id', '')}\n"
                                    f"Nombre d'échantillons : {sample_count}\n"
                                    f"Analyses demandées :\n{analysis_details_fr}\n"
                                    f"Type de laboratoire : {visibility}\n"
                                    f"Opérateur en charge de l’analyse : {record.get('assigned_operator') or 'Non renseigné'}\\n"
                                    f"{('Prix total indicatif : ' + price_text_fr) if visibility == 'Public (IJPB)' else price_text_fr}\n"
                                    f"{quote_followup_fr}\n"
                                    f"{disclaimer_fr}\n\n"
                                    f"Commentaire : {comment.strip() or 'Aucun commentaire.'}\n\n"
                                    "Cordialement,\nPlateforme OV-Chimie-Métabolisme"
                                )
                            else:
                                body = (
                                    f"Hello {record.get('project_manager_first_name') or record.get('project_manager', '')},\n\n"
                                    "Your analysis request has been accepted.\n\n"
                                    f"Request reference: {record.get('id', '')}\n"
                                    f"Number of samples: {sample_count}\n"
                                    f"Requested analyses:\n{analysis_details_en}\n"
                                    f"Laboratory type: {visibility}\n"
                                    f"Operator responsible for the analysis: {record.get('assigned_operator') or 'Not specified'}\\n"
                                    f"{('Indicative total price: ' + price_text_en) if visibility == 'Public (IJPB)' else price_text_en}\n"
                                    f"{quote_followup_en}\n"
                                    f"{disclaimer_en}\n\n"
                                    f"Comment: {comment.strip() or 'No comment.'}\n\n"
                                    "Best regards,\nOV-Chimie-Metabolisme Platform"
                                )
                        else:
                            body = t("mail_user_body").format(
                                manager=record.get("project_manager", ""),
                                project=record.get("project_name", ""),
                                status=status_display[status_options.index(new_status)],
                                comment=comment.strip() or ("Aucun commentaire." if st.session_state.language == "fr" else "No comment."),
                            )
                        mail_subject = (
                            f"Demande d'analyse acceptée – Réf. {record.get('id', '')}"
                            if new_status == "accepted" and st.session_state.language == "fr"
                            else f"Analysis request accepted – Ref. {record.get('id', '')}"
                            if new_status == "accepted"
                            else t("mail_user_subject")
                        )
                        send_email(
                            email, mail_subject, body,
                            cc=join_email_recipients(*cc_addresses),
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
                f"{request_display_label(r)} — {r.get('id', '')}": r["id"]
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
    steps = t("steps")
    step = min(st.session_state.step, len(steps) - 1)

    hero_title = (
        "OV-Chimie-Métabolisme/IJPB-INRAE Versailles"
        if st.session_state.get("language", "fr") == "fr"
        else "PO Chemistry-Metabolism/IJPB-INRAE Versailles"
    )
    hero_html = """
        <style>
        .block-container {
            max-width: none !important;
            width: 100% !important;
            margin-left: 0 !important;
            margin-right: 0 !important;
            padding-left: 1.25rem !important;
            padding-right: 1.25rem !important;
            padding-top: 1.25rem;
            padding-bottom: 2rem;
        }
        .hero {padding: 1.3rem 1.5rem; border-radius: 14px;
               background: linear-gradient(120deg,#edf5fb,#f4f7f2); margin-bottom: 1rem;}
        .hero h1 {margin: 0; color: #183b56; font-size: 1.8rem;}
        .hero p {margin: .4rem 0 0; color: #526575;}
        </style>
        <div class="hero"><h1>__HERO_TITLE__</h1>
        <p>PO-Chem-WebApp-Request</p></div>
        """.replace("__HERO_TITLE__", hero_title)
    st.markdown(hero_html, unsafe_allow_html=True)
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
                text_field("project_manager", t("manager_last_name"))
                text_field("project_manager_first_name", t("manager_first_name"))
                text_field("group", t("group"))
                selected_date = st.date_input(t("date"), value=date.fromisoformat(data["date"]) if data.get("date") else date.today())
                data["date"] = selected_date.isoformat()
                text_field("manager_phone", t("phone"))
                text_field("manager_email", t("manager_email"))
                visibility_values = ["Privée", "Public", "Public (IJPB)"]
                visibility_labels = [
                    t("visibility_private"), t("visibility_public"), t("visibility_ijpb")
                ]
                saved_visibility = data.get("visibility", "Public (IJPB)")
                visibility_index = visibility_values.index(saved_visibility) if saved_visibility in visibility_values else 0
                selected_visibility = st.radio(
                    t("visibility"),
                    options=visibility_labels,
                    index=visibility_index,
                    key="request_visibility",
                )
                data["visibility"] = visibility_values[visibility_labels.index(selected_visibility)]
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
            # One request can contain several analyses, each with its own sample count and mode.
            analysis_items = data.setdefault("analysis_items", [])
            if not analysis_items:
                # Keep compatibility with an already populated legacy single-analysis form,
                # but start a genuinely new request with empty fields.
                legacy_type = data.get("analysis_type", "")
                analysis_items.append({
                    "analysis_type": legacy_type,
                    "analysis_mode": data.get("analysis_mode", "") if legacy_type else "",
                    "sample_count": data.get("sample_count") if legacy_type else None,
                    "analysis_details": data.get("analysis_details", ""),
                })
            st.markdown(f"**{t('analysis_type')}**")
            for analysis_index, analysis_item in enumerate(analysis_items):
                current = analysis_item.get("analysis_type", "") or ""
                summary_type = (current if st.session_state.language == "fr" else available_analysis_type_translations.get(current, ANALYSIS_TYPES_EN.get(current, current))) if current else ("À renseigner" if st.session_state.language == "fr" else "To be completed")
                raw_samples = analysis_item.get("sample_count")
                summary_samples = f"{raw_samples} échantillon(s)" if raw_samples not in (None, "") else "nombre à renseigner" if st.session_state.language == "fr" else "sample count to be entered"
                expander_label = (f"Analyse {analysis_index + 1} · {summary_type} · {summary_samples}" if st.session_state.language == "fr" else f"Analysis {analysis_index + 1} · {summary_type} · {summary_samples}")
                with st.expander(expander_label, expanded=False):
                    labels, values = [], []
                    for canonical_type in available_analysis_types:
                        translated = canonical_type if st.session_state.language == "fr" else available_analysis_type_translations.get(canonical_type, ANALYSIS_TYPES_EN.get(canonical_type, canonical_type))
                        labels.append(translated); values.append(canonical_type)
                    current_index = values.index(current) if current in values else None
                    selected_label = st.selectbox(
                        t("analysis_type"), labels, index=current_index,
                        placeholder="Sélectionner un type d’analyse" if st.session_state.language == "fr" else "Select an analysis type",
                        key=f"analysis_type_{analysis_index}"
                    )
                    selected_type = values[labels.index(selected_label)] if selected_label else ""
                    analysis_item["analysis_type"] = selected_type
                    atype = selected_type
                    modes = ["ESI+", "ESI-", "ESI+/-"] if atype.endswith("LC-HRMS") else (["Splitless", "Split", "Split/Splitless"] if atype.endswith("GC-MS") else [])
                    if modes:
                        previous = analysis_item.get("analysis_mode")
                        mode_index = modes.index(previous) if previous in modes else None
                        mode_label = t("analysis_mode_lchrms") if atype.endswith("LC-HRMS") else t("analysis_mode_gc_ms")
                        analysis_item["analysis_mode"] = st.selectbox(
                            mode_label, modes, index=mode_index,
                            placeholder="Sélectionner un mode" if st.session_state.language == "fr" else "Select a mode",
                            key=f"analysis_mode_{analysis_index}"
                        ) or ""
                    else:
                        analysis_item["analysis_mode"] = ""
                    previous_count = analysis_item.get("sample_count")
                    analysis_item["sample_count"] = st.number_input(
                        t("sample_count"), min_value=1, max_value=1000000,
                        value=int(previous_count) if previous_count not in (None, "") else None,
                        step=1, key=f"analysis_samples_{analysis_index}"
                    )
                    if atype.startswith(("Développement de méthodes spécifiques", "Développements spécifiques")):
                        analysis_item["analysis_details"] = st.text_area(t("specific_development_details"), value=analysis_item.get("analysis_details", ""), height=100, key=f"analysis_details_{analysis_index}")
                    else:
                        analysis_item["analysis_details"] = ""
                    if len(analysis_items) > 1 and st.button(("Supprimer cette analyse" if st.session_state.language == "fr" else "Remove this analysis"), key=f"remove_analysis_{analysis_index}"):
                        analysis_items.pop(analysis_index); st.rerun()
            if st.button(("＋ Ajouter un autre type d’analyse" if st.session_state.language == "fr" else "＋ Add another analysis type"), key="add_analysis_item"):
                analysis_items.append({"analysis_type": "", "analysis_mode": "", "sample_count": None, "analysis_details": ""})
                st.rerun()
            data["analysis_items"] = analysis_items
            data["analysis_type"] = analysis_items[0].get("analysis_type", "")
            data["analysis_mode"] = analysis_items[0].get("analysis_mode", "")
            data["analysis_details"] = analysis_items[0].get("analysis_details", "")
            data["sample_count"] = sum(int(x.get("sample_count") or 0) for x in analysis_items)
            text_field("platform_staff_email", t("platform_staff_email"))
            st.caption(t("recommended"))
        elif step == 1:
            st.subheader(t("description"))
            st.caption(t("description_help"))
            text_area("project_description", t("description_field"), 230)
        elif step == 2:
            st.subheader(t("samples_title"))
            st.caption(t("samples_help"))
            st.caption("Le nombre d’échantillons est renseigné séparément pour chaque analyse dans la rubrique Plateforme." if st.session_state.language == "fr" else "The sample count is entered separately for each analysis in the Platform section.")
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
                (t("manager"), " ".join(part for part in (data.get("project_manager", ""), data.get("project_manager_first_name", "")) if part)),
                (t("visibility"), data.get("visibility", "Public (IJPB)")),
                (t("sample_count"), str(data.get("sample_count", ""))),
                (t("group"), data.get("group", "")),
                (t("date"), data.get("date", "")),
                (t("manager_email"), data.get("manager_email", "")),
                (t("analysis_type"), " ; ".join(f"{x.get('analysis_type', '')} — {x.get('sample_count', 1)}" for x in data.get("analysis_items", []))),
                (t("project_name"), data.get("project_name", "")),
                (t("funding"), data.get("funding", "NO")),
                (t("guidance"), data.get("guidance", "NO")),
            ]
            for label, value in summary:
                st.markdown(f"**{label} :** {value or '—'}")

            charter_accepted = st.checkbox(
                "J’ai lu et j’accepte la charte de la plateforme."
                if st.session_state.language == "fr"
                else "I have read and agree to the platform charter.",
                value=False,
                key="charter_accepted_checkbox",
            )
            st.caption(
                "La charte est disponible dans le menu latéral, sous le logo de la plateforme, "
                "en cliquant sur « Télécharger la charte de la plate-forme (PDF) »."
                if st.session_state.language == "fr"
                else "The charter is available in the sidebar, below the platform logo, "
                "using the “Download the platform charter (PDF)” button."
            )

            with st.expander(t("all_answers")):
                answer_labels = {
                    "project_manager": t("manager_last_name"),
                    "project_manager_first_name": t("manager_first_name"),
                    "visibility": t("visibility"),
                    "group": t("project_group_short"),
                    "date": t("project_date_short"),
                    "manager_phone": t("phone"),
                    "manager_email": t("project_manager_email_short"),
                    "admin_staff": t("project_staff_short"),
                    "admin_phone": t("project_staff_phone_short"),
                    "admin_email": t("project_staff_email_short"),
                    "platform_staff_email": t("platform_staff_email_short"),
                    "analysis_type": t("analysis_type"),
                    "analysis_mode": t("analysis_mode"),
                    "analysis_details": t("specific_development_details"),
                    "project_description": t("project_description_short"),
                    "samples": t("samples_short"),
                    "sample_count": t("sample_count"),
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
                data.get("project_manager_first_name", "").strip(),
                data.get("manager_email", "").strip(),
                data.get("group", "").strip(),
                data.get("project_name", "").strip(),
                data.get("analysis_type", "").strip(),
            ]
            analyses_valid = bool(data.get("analysis_items")) and all(str(x.get("analysis_type", "")).strip() and int(x.get("sample_count", 0) or 0) > 0 and (not str(x.get("analysis_type", "")).startswith(("Développement de méthodes spécifiques", "Développements spécifiques")) or str(x.get("analysis_details", "")).strip()) for x in data.get("analysis_items", []))
            email_is_valid = bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", data.get("manager_email", "").strip()))
            specific_details_valid = (
                not (data.get("analysis_type", "").startswith("Développement de méthodes spécifiques") or data.get("analysis_type", "").startswith("Développements spécifiques"))
                or bool(data.get("analysis_details", "").strip())
            )
            form_is_valid = all(required_values) and email_is_valid and specific_details_valid and analyses_valid and charter_accepted
            if not form_is_valid:
                if not charter_accepted:
                    st.warning(
                        "Vous devez lire et accepter la charte de la plateforme avant de soumettre la demande."
                        if st.session_state.language == "fr"
                        else "You must read and accept the platform charter before submitting the request."
                    )
                elif (data.get("analysis_type", "").startswith("Développement de méthodes spécifiques") or data.get("analysis_type", "").startswith("Développements spécifiques")) and not data.get("analysis_details", "").strip():
                    st.warning(t("specific_development_required"))
                else:
                    st.warning(t("required"))
            else:
                if not data.get("request_id"):
                    request_timestamp = datetime.now().astimezone().strftime("%Y%m%d-%H%M%S")
                    request_identifier = uuid.uuid4().hex[:8].upper()
                    data["request_id"] = f"REQ-{request_identifier}-{request_timestamp}"
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
                                "visibility": data.get("visibility", "Public (IJPB)"),
                                "project_manager": data.get("project_manager", "").strip(),
                                "project_manager_first_name": data.get("project_manager_first_name", "").strip(),
                                "manager_email": data.get("manager_email", ""),
                                "admin_staff": data.get("admin_staff", ""),
                                "admin_phone": data.get("admin_phone", ""),
                                "admin_email": data.get("admin_email", ""),
                                "platform_staff_email": data.get("platform_staff_email", ""),
                                "analysis_type": data.get("analysis_type", ""),
                                "analysis_mode": data.get("analysis_mode", ""),
                                "analysis_details": data.get("analysis_details", ""),
                                "sample_count": int(data.get("sample_count", 1) or 1),
                                "analysis_items": [{**x, "sample_count": int(x.get("sample_count", 1) or 1)} for x in data.get("analysis_items", [])],
                                "group": data.get("group", ""),
                                "request_date": data.get("date", ""),
                                "submitted_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                                "status": "pending",
                                "admin_comment": "",
                                "charter_accepted": True,
                                "charter_accepted_at": datetime.now().astimezone().isoformat(timespec="seconds"),
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
                                    recipients.get("ijpb_admin_email", ""),
                                    record.get("admin_email", ""),
                                    record.get("platform_staff_email", ""),
                                )
                                send_email(
                                    admin_email,
                                    t("mail_admin_subject"),
                                    t("mail_admin_body").format(
                                        project=record["project_name"],
                                        manager=record.get("project_manager_first_name") or record["project_manager"],
                                        group=record["group"],
                                        analysis_type=" ; ".join(x.get("analysis_type", "") for x in record.get("analysis_items", [])) or record.get("analysis_type", ""),
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

    # Quick navigation to any of the six request sections.
    st.markdown(
        "**Accès rapide aux rubriques**" if st.session_state.language == "fr"
        else "**Quick section navigation**"
    )
    step_columns = st.columns(len(steps))
    for step_index, step_label in enumerate(steps):
        with step_columns[step_index]:
            button_label = step_label
            if st.button(
                button_label,
                key=f"quick_step_navigation_{step_index}",
                use_container_width=True,
                type="primary" if step_index == step else "secondary",
            ):
                st.session_state.step = step_index
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

# Global language selector: available throughout the application and
# deliberately independent from the FAIR user-profile preferences.
if "language_selector" not in st.session_state:
    st.session_state.language_selector = (
        "Français" if st.session_state.language == "fr" else "English"
    )
_selected_language_label = st.sidebar.selectbox(
    "Langue / Language",
    options=["Français", "English"],
    key="language_selector",
)
_selected_language_code = "fr" if _selected_language_label == "Français" else "en"
if st.session_state.language != _selected_language_code:
    st.session_state.language = _selected_language_code
    st.rerun()

# Profiles are loaded and displayed only inside FAIR > Analysis.
profile_payload = {"profiles": {}}
profile_storage_ok = False
profile_names = []



def render_keep_preferences_button(sheet):
    """Save the current sheet's interface preferences to the active Nextcloud profile."""
    fr = st.session_state.get("language", "fr") == "fr"
    label = "Keep as preference in my profile" if not fr else "Garder comme préférences dans mon profil"
    if st.button(label, key=f"keep_preferences_{sheet.lower().replace(' ', '_')}", use_container_width=True):
        try:
            payload = load_user_profiles()
            active_name = st.session_state.get("active_user_profile")
            profiles = payload.get("profiles", {})
            if not active_name or active_name not in profiles:
                st.error("Aucun profil actif n'est disponible." if fr else "No active profile is available.")
                return
            profile = profiles[active_name]
            preferences_by_sheet = profile.setdefault("preferences_by_sheet", {})
            preferences_by_sheet[sheet] = get_profile_preferences(sheet)
            save_user_profiles(payload)
            st.success(
                f"Préférences de {sheet} enregistrées dans le profil « {active_name} »."
                if fr else
                f"{sheet} preferences saved to profile '{active_name}'."
            )
        except Exception as exc:
            st.error(
                ("Impossible d'enregistrer les préférences : " if fr else "Could not save preferences: ")
                + str(exc)
            )

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
    with st.expander("Enregistrer sous un nouveau profil" if fr else "Save as a new profile"):
        st.caption(
            "Crée une copie de tes préférences actuelles pour les trois feuilles : Analysis, Data Treatment et Pipeline Step by Step. Les données FAIR ne sont jamais copiées."
            if fr else
            "Create a copy of your current preferences for all three sheets: Analysis, Data Treatment and Pipeline Step by Step. FAIR data are never copied."
        )
        with st.form(f"create_user_profile_form_{sheet.lower().replace(' ', '_')}", clear_on_submit=True):
            new_profile_name = st.text_input(
                "Nom du nouveau profil" if fr else "New profile name",
                placeholder="Ex. Marie, Jean, Chriso…" if fr else "e.g. Marie, Jean, Chris…",
            )
            create_profile_submitted = st.form_submit_button(
                "Enregistrer sous" if fr else "Save as", use_container_width=True
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
                import copy
                active_data = profile_payload["profiles"].get(st.session_state.active_user_profile, {})
                copied_preferences = copy.deepcopy(active_data.get("preferences_by_sheet", {}))
                if "preferences" in active_data and "Analysis" not in copied_preferences:
                    copied_preferences["Analysis"] = copy.deepcopy(active_data["preferences"])
                copied_preferences[sheet] = get_profile_preferences(sheet)
                profile_payload["profiles"][clean_name] = {"preferences_by_sheet": copied_preferences}
                try:
                    save_user_profiles(profile_payload)
                    st.session_state.active_user_profile = clean_name
                    st.session_state.pop(loaded_marker, None)
                    st.rerun()
                except Exception as exc:
                    profile_payload["profiles"].pop(clean_name, None)
                    st.error(("Impossible de créer le profil : " if fr else "Could not create profile: ") + str(exc))

    with st.expander("Ajouter un profil" if fr else "Add a profil"):
        st.caption(
            "Le nouvel utilisateur reprend les préférences du dernier profil actif pour Analysis, Data Treatment et Pipeline Step by Step. Chaque utilisateur pourra ensuite modifier ses propres préférences."
            if fr else
            "The new user inherits the preferences of the last active profile for Analysis, Data Treatment and Pipeline Step by Step. Each user can then customize their own preferences."
        )
        with st.form(f"add_user_profile_form_{sheet.lower().replace(' ', '_')}", clear_on_submit=True):
            added_user_name = st.text_input(
                "Nom du nouvel utilisateur" if fr else "New user name",
                placeholder="Ex. Marie, Jean…" if fr else "e.g. Marie, Jean…",
            )
            add_user_submitted = st.form_submit_button(
                "Ajouter l'utilisateur" if fr else "Add user", use_container_width=True
            )
        if add_user_submitted:
            clean_user_name = added_user_name.strip()
            if not clean_user_name:
                st.warning("Saisis un nom d'utilisateur." if fr else "Enter a user name.")
            elif any(existing.casefold() == clean_user_name.casefold() for existing in profile_names):
                st.warning("Cet utilisateur existe déjà." if fr else "This user already exists.")
            elif not profile_storage_ok:
                st.error("Le stockage Nextcloud des profils n'est pas disponible." if fr else "Nextcloud profile storage is unavailable.")
            else:
                import copy
                # Snapshot all preferences from the currently active (last active) user.
                source_data = profile_payload["profiles"].get(st.session_state.active_user_profile, {})
                inherited = copy.deepcopy(source_data.get("preferences_by_sheet", {}))
                if "preferences" in source_data and "Analysis" not in inherited:
                    inherited["Analysis"] = copy.deepcopy(source_data["preferences"])
                inherited[sheet] = get_profile_preferences(sheet)
                profile_payload["profiles"][clean_user_name] = {"preferences_by_sheet": inherited}
                try:
                    save_user_profiles(profile_payload)
                    st.session_state.active_user_profile = clean_user_name
                    st.session_state.pop(loaded_marker, None)
                    st.rerun()
                except Exception as exc:
                    profile_payload["profiles"].pop(clean_user_name, None)
                    st.error(("Impossible d'ajouter l'utilisateur : " if fr else "Could not add user: ") + str(exc))

    with st.expander("Renommer le profil actif" if fr else "Rename active profile"):
        with st.form(f"rename_user_profile_form_{sheet.lower().replace(' ', '_')}"):
            renamed_profile = st.text_input(
                "Nouveau nom" if fr else "New name",
                value=st.session_state.active_user_profile,
            )
            rename_submitted = st.form_submit_button("Renommer" if fr else "Rename", use_container_width=True)
        if rename_submitted:
            clean_name = renamed_profile.strip()
            old_name = st.session_state.active_user_profile
            if not clean_name:
                st.warning("Saisis un nom." if fr else "Enter a name.")
            elif clean_name.casefold() != old_name.casefold() and any(n.casefold() == clean_name.casefold() for n in profile_names):
                st.warning("Ce nom de profil existe déjà." if fr else "That profile name already exists.")
            elif not profile_storage_ok:
                st.error("Le stockage Nextcloud des profils n'est pas disponible." if fr else "Nextcloud profile storage is unavailable.")
            else:
                profile_payload["profiles"][clean_name] = profile_payload["profiles"].pop(old_name)
                try:
                    save_user_profiles(profile_payload)
                    st.session_state.active_user_profile = clean_name
                    st.session_state.pop(loaded_marker, None)
                    st.rerun()
                except Exception as exc:
                    profile_payload["profiles"][old_name] = profile_payload["profiles"].pop(clean_name)
                    st.error(("Impossible de renommer le profil : " if fr else "Could not rename profile: ") + str(exc))

    with st.expander("Supprimer un profil" if fr else "Delete a profile"):
        if len(profile_names) <= 1:
            st.caption("Il faut conserver au moins un profil." if fr else "At least one profile must be kept.")
        else:
            st.warning(
                f"Le profil actif « {st.session_state.active_user_profile} » et ses préférences seront supprimés définitivement."
                if fr else
                f"The active profile '{st.session_state.active_user_profile}' and its preferences will be permanently deleted."
            )
            with st.form(f"delete_user_profile_form_{sheet.lower().replace(' ', '_')}"):
                confirm_delete = st.checkbox("Je confirme la suppression" if fr else "I confirm deletion")
                delete_submitted = st.form_submit_button("Supprimer le profil actif" if fr else "Delete active profile", use_container_width=True)
            if delete_submitted:
                if not confirm_delete:
                    st.warning("Coche la confirmation avant de supprimer." if fr else "Check the confirmation box before deleting.")
                elif not profile_storage_ok:
                    st.error("Le stockage Nextcloud des profils n'est pas disponible." if fr else "Nextcloud profile storage is unavailable.")
                else:
                    old_name = st.session_state.active_user_profile
                    removed = profile_payload["profiles"].pop(old_name, None)
                    try:
                        save_user_profiles(profile_payload)
                        remaining = sorted(profile_payload["profiles"].keys(), key=str.casefold)
                        st.session_state.active_user_profile = remaining[0]
                        st.session_state.pop(loaded_marker, None)
                        st.rerun()
                    except Exception as exc:
                        if removed is not None:
                            profile_payload["profiles"][old_name] = removed
                        st.error(("Impossible de supprimer le profil : " if fr else "Could not delete profile: ") + str(exc))

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
        ("Contact" if _fr_navigation else "Contact"): "contact",
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


# Contact is a dedicated main page so the form uses the full content area.
_contact_fr = st.session_state.get("language", "fr") == "fr"


# Shared platform resources: always visible below the navigation menu for every user.
_footer_fr = st.session_state.get("language", "fr") == "fr"
st.sidebar.divider()
# One coherent banner containing the OV-Chimie logo and the partner logos.
# Keep this asset at the application root so it is included in simple ZIP deployments.
_platform_banner_path = os.path.join(os.path.dirname(__file__), "ov_chimie_logo_user.png")
if os.path.isfile(_platform_banner_path):
    st.sidebar.image(_platform_banner_path, use_container_width=True)
else:
    st.sidebar.warning(
        "Logo de la plate-forme introuvable." if _footer_fr
        else "Platform logo image not found."
    )
st.sidebar.markdown(
    "[Site web de la plate-forme](https://green-chemistry-platform.versailles-saclay.hub.inrae.fr/)"
    if _footer_fr else
    "[Platform website](https://green-chemistry-platform.versailles-saclay.hub.inrae.fr/)"
)
_charter_path = os.path.join(os.path.dirname(__file__), "Charte E-GES-00-ver4.pdf")
if os.path.isfile(_charter_path):
    with open(_charter_path, "rb") as _charter_file:
        st.sidebar.download_button(
            label="Télécharger la charte de la plate-forme (PDF)" if _footer_fr else "Download the platform charter (PDF)",
            data=_charter_file.read(),
            file_name="Charte_E-GES-00_version_4.pdf",
            mime="application/pdf",
            key="download_platform_charter_sidebar",
            use_container_width=True,
        )

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
elif page == "contact":
    st.title("Contacter l’administrateur" if st.session_state.language == "fr" else "Contact the administrator")
    st.write(
        "Utilisez ce formulaire pour envoyer directement un message à l’administrateur de la plateforme."
        if st.session_state.language == "fr" else
        "Use this form to send a message directly to the platform administrator."
    )
    with st.form("main_contact_form", clear_on_submit=True):
        _contact_sender = st.text_input(
            "Votre adresse e-mail *" if _contact_fr else "Your email address *",
            placeholder="nom@exemple.fr",
        )
        _contact_subject = st.text_input(
            "Objet *" if _contact_fr else "Subject *",
        )
        _contact_message = st.text_area(
            "Message *" if _contact_fr else "Message *",
            height=300,
            placeholder=("Écrivez votre message ici…" if _contact_fr else "Write your message here…"),
        )
        _contact_submit = st.form_submit_button(
            "Envoyer le message" if _contact_fr else "Send message",
            use_container_width=True,
        )
    if _contact_submit:
        _email_pattern = r"^[^\s@]+@[^\s@]+\.[^\s@]+$"
        if not re.match(_email_pattern, _contact_sender.strip()):
            st.error("Veuillez saisir une adresse e-mail valide." if _contact_fr else "Please enter a valid email address.")
        elif not _contact_subject.strip() or not _contact_message.strip():
            st.error("Veuillez renseigner l’objet et le message." if _contact_fr else "Please enter both a subject and a message.")
        else:
            try:
                _contact_settings = load_notification_settings()
                _contact_admin_email = str(_contact_settings.get("admin_email") or DEFAULT_ADMIN_EMAIL).strip()
                if not _contact_admin_email:
                    raise RuntimeError("L'adresse e-mail principale de l'administrateur n'est pas configurée." if _contact_fr else "The administrator's primary email address is not configured.")
                _contact_mail_subject = ("[Contact plateforme] " if _contact_fr else "[Platform contact] ") + _contact_subject.strip()
                _contact_mail_body = (
                    f"Message envoyé depuis le formulaire Contact de PO-Chem-WebApp.\n\n"
                    f"Adresse de réponse : {_contact_sender.strip()}\n"
                    f"Objet : {_contact_subject.strip()}\n\n"
                    f"Message :\n{_contact_message.strip()}"
                )
                send_email(
                    _contact_admin_email,
                    _contact_mail_subject,
                    _contact_mail_body,
                    reply_to=_contact_sender.strip(),
                )
                st.success("Votre message a été envoyé à l’administrateur." if _contact_fr else "Your message has been sent to the administrator.")
            except Exception as _contact_exc:
                st.error(("L’envoi du message a échoué : " if _contact_fr else "Message could not be sent: ") + str(_contact_exc))
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
