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
