PO-Chem-WebApp-Request — Version 1.4.7

Fichiers à placer à la racine du dépôt GitHub :
- app.py
- requirements.txt
- template_E-GES-01.docx

Nouveauté — Types d'analyse entièrement bilingues :
- Lorsqu'il ajoute un type, l'administrateur saisit le libellé français et sa traduction anglaise.
- L'administrateur peut également sélectionner un type existant et modifier sa traduction anglaise.
- Les traductions personnalisées sont sauvegardées dans MyDrive/notification_settings.json sur Nextcloud.
- Le formulaire public affiche le libellé correspondant à la langue choisie (FR/EN).
- Les statistiques et les détails de demande dans l'espace administrateur utilisent également la traduction anglaise lorsque l'interface est en anglais.
- Les anciens réglages Nextcloud restent compatibles : les types prédéfinis conservent leurs traductions ; les types personnalisés déjà créés peuvent recevoir leur traduction via « Modifier une traduction anglaise ».
- Les types supprimés de la liste restent présents dans l'historique des anciennes demandes.

Déploiement :
Remplacer app.py et README_DEPLOIEMENT.txt dans le dépôt GitHub. Les autres fichiers restent identiques.
Conserver les secrets Streamlit existants.

Module FAIR – Sample (première version)
- Un classeur FAIR est créé à chaque nouvelle demande et stocké dans le même dossier Nextcloud que le document Word, avec le même nom suivi de _FAIR.xlsx.
- Le classeur source FAIR_template.xlsx doit être conservé à la racine du dépôt avec app.py. Il contient les feuilles Sample, Analysis, Data treatment et Pipeline Step by Step ; seule la feuille Sample est alimentée dans cette version.
- La feuille Sample conserve son orientation (métadonnées en lignes, échantillons en colonnes) et sa mise en forme. Les exemples de données sont vidés à la création ; les champs de demande sont préremplis automatiquement.
- Le demandeur dispose de quatre modes : formulaire web, import Excel/CSV, tableau interactif et téléchargement du modèle FAIR.
- Les données sont rapprochées avec ID tube analyse. Les champs File ID Q-Tof ESI+ et File ID Q-Tof ESI- sont réservés à l'administrateur. Les imports du demandeur ne peuvent pas les modifier.
- Dans l'espace administrateur, le bouton « Demander le remplissage FAIR » envoie le lien sécurisé au demandeur. L'envoi nécessite le secret Streamlit APP_BASE_URL, défini à l'URL publique de l'application (ex. https://mon-app.streamlit.app).
- L'administrateur peut importer les File IDs à partir du modèle administrateur ; les IDs inconnus sont refusés.
- La suppression définitive d'une demande supprime aussi son classeur FAIR associé.
- Les anciennes demandes, créées avant ce module, peuvent utiliser « Créer / vérifier le classeur FAIR » pour créer leur classeur à partir du modèle.
