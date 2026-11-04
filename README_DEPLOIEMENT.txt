# PO-Chem-WebApp-Request — E-GES-01

## Déploiement
Repository : `TOTOZAFY/PO-Chem-WebApp-request-IJPB`
Branch : `main`
Main file : `app.py`

Fichiers à la racine :
- `app.py`
- `requirements.txt`
- `template_E-GES-01.docx`

## Secrets Streamlit

Dans Streamlit Community Cloud > application > Settings > Secrets, configurer :

```toml
NEXTCLOUD_URL = "https://nextcloud.inrae.fr"
NEXTCLOUD_USERNAME = "VOTRE_IDENTIFIANT_NEXTCLOUD"
NEXTCLOUD_APP_PASSWORD = "VOTRE_MOT_DE_PASSE_APPLICATION"
ADMIN_PASSWORD = "CHOISIR_UN_MOT_DE_PASSE_ADMIN_FORT"
# Facultatif : adresse mise en copie des notifications de nouvelles demandes
ADMIN_CC_EMAIL = "suivi@exemple.fr"

[smtp]
host = "SERVEUR_SMTP_INRAE"
port = 587
username = "COMPTE_SMTP"
password = "MOT_DE_PASSE_SMTP"
sender = "ADRESSE_EXPEDITRICE_AUTORISEE"
```

Ne jamais inscrire les mots de passe dans le code ni les communiquer dans une conversation.
Le serveur SMTP et l'autorisation d'envoi doivent être fournis/validés par l'INRAE.
L'adresse destinataire principale des nouvelles demandes est fixée à `ov-chimie@inrae.fr`.
`ADMIN_CC_EMAIL` est facultative : si elle est renseignée, cette adresse reçoit une copie (CC) des notifications de nouvelles demandes pour le suivi interne. Si elle est vide ou absente, aucun CC n'est ajouté.

## Fonctionnement
- Le formulaire génère le document Word à partir du modèle fourni.
- Le document est déposé dans `MyDrive/E-GES-01/Demandes_Word/`.
- Les métadonnées (projet, responsable, groupe, date, statut, commentaire et chemin du document) sont stockées dans `MyDrive/E-GES-01/demandes_metadata.json`.
- Une notification est envoyée à `ov-chimie@inrae.fr` après la soumission.
- L'espace administrateur est protégé par `ADMIN_PASSWORD`.
- L'administrateur peut choisir En attente / Acceptée / Refusée et ajouter un commentaire. La décision est enregistrée puis notifiée par e-mail au demandeur.
- Le document Word n'est pas joint aux e-mails.

## Important
1. Cette version utilise WebDAV de Nextcloud. L'URL de base est transformée en endpoint WebDAV `/remote.php/dav/files/<identifiant>/`.
2. Vérifier que l'espace `MyDrive` est bien accessible en WebDAV avec le compte concerné. Les règles d'accès INRAE peuvent restreindre WebDAV ou SMTP.
3. Le stockage JSON convient à un volume modéré et à un usage administrateur unique. Pour des usages concurrents ou institutionnels à grande échelle, préférer une base de données validée par l'INRAE.
4. Faire valider par l'INRAE l'hébergement Streamlit Cloud, le traitement des données personnelles et les notifications avant toute utilisation en production.
5. La page 2 du modèle Word, réservée aux responsables de plateforme, n'est pas modifiée par le code.


CONFIGURATION DES NOTIFICATIONS (version 1.3)
------------------------------------------------
Les trois adresses e-mail sont maintenant modifiables depuis l'espace administrateur :
- E-mail principal de l'administrateur : destinataire des nouvelles demandes.
- E-mail secondaire de l'administrateur (CC, facultatif) : copie des nouvelles demandes.
- E-mail de suivi du personnel (facultatif) : reçoit les nouvelles demandes ainsi que les décisions d'acceptation ou de refus.

Les paramètres sont enregistrés dans Nextcloud, dans :
MyDrive/E-GES-01/notification_settings.json

Le premier accès utilise par défaut ov-chimie@inrae.fr comme adresse principale.
Le personnel de suivi est mis en copie des notifications de décision adressées au demandeur.


MISE À JOUR DU SUIVI DU PERSONNEL (version 1.3.1)
--------------------------------------------------
Le suivi du personnel est désormais propre à chaque demande :
- L'adresse du champ « E-mail du personnel administratif » du formulaire reçoit la notification liée à cette demande.
- Elle reçoit également la notification lorsqu'une décision (acceptation/refus) est prise pour cette demande.
- Aucun autre personnel ne reçoit ces notifications.
- L'administrateur principal et son adresse secondaire en CC restent configurables dans l'espace administrateur.
