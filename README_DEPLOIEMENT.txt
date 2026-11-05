PO-Chem-WebApp-Request — Version 1.4.7

Fichiers à placer à la racine du dépôt GitHub :
- app.py
- requirements.txt
- template_E-GES-01.docx
- FAIR_template.xlsx

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


Nouveauté — Menu principal FAIR
- Un menu FAIR est maintenant visible dans la navigation principale, côté demandeur et administrateur.
- Le demandeur saisit la référence de sa demande et l'adresse e-mail utilisée lors de la soumission. La combinaison des deux identifie sa demande et ouvre uniquement son classeur FAIR.
- L'administrateur authentifié peut sélectionner une demande directement depuis le menu FAIR.
- Le menu présente les quatre feuilles du classeur : Sample, Analysis, Data treatment et Pipeline Step by Step. Seule la feuille Sample est modifiable à ce stade ; les trois autres sont affichées comme à venir.
- L'e-mail de demande de remplissage FAIR indique la référence de la demande, explique comment accéder au menu FAIR et conserve également le lien direct sécurisé.

Mise à jour — Menu FAIR avec choix du rôle
- Le menu principal FAIR commence par un choix explicite Demandeur / Administrateur.
- Dans Demandeur, l'utilisateur saisit la référence de sa demande et son adresse e-mail ; il ne voit que les champs qui lui sont destinés dans Sample (Plant Collection et métadonnées des échantillons). Les File IDs administrateur ne sont jamais affichés ni modifiables.
- Dans Administrateur, une authentification par ADMIN_PASSWORD est demandée dans le menu FAIR. L'administrateur sélectionne ensuite une demande et ne voit que les champs qui lui sont réservés dans Sample (ID tube analyse et File ID Q-Tof ESI+ / ESI-), avec édition en tableau et import Excel/CSV.
- Les feuilles Analysis, Data treatment et Pipeline Step by Step restent visibles comme sections à venir. Leur contenu sera développé ultérieurement.
- Les deux rôles modifient le même classeur FAIR sur Nextcloud ; la fusion est effectuée par ID tube analyse et les champs de l'autre rôle sont préservés.


Version 1.4 - Navigation FAIR : sélectionner d'abord la feuille (Sample, Analysis, Data treatment, Pipeline Step by Step), puis dans Sample choisir Demandeur ou Administrateur. Les deux rôles n'affichent que leurs propres champs.
