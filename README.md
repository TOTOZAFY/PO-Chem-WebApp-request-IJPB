# PO-Chem-WebApp-request-IJPB
Web application for project requests


## FAIR navigation

The FAIR menu is organized by workbook sheet first: Sample, Analysis, Data treatment, and Pipeline Step by Step. Inside Sample, users choose Requester or Administrator. Each role sees only its assigned fields. The three other sheets are visible as future sections and are not yet editable.


## Profils FAIR

Le sélecteur de profil et la création de profils sont disponibles dans les feuilles FAIR `Analysis`, `Data treatment` et `Pipeline Step by Step`. Chaque profil conserve un espace de préférences distinct pour chacune de ces feuilles dans `MyDrive/user_profiles.json` sur Nextcloud. La feuille `Sample` et les autres rubriques de l'application ne chargent ni n'enregistrent de préférences de profil. Les anciens profils contenant le format v1.7-v1.9 (`preferences`) sont migrés automatiquement vers `preferences_by_sheet` en conservant leurs préférences `Analysis`. Les données FAIR et les réponses des demandes ne sont pas stockées dans les profils.

`Data treatment` et `Pipeline Step by Step` restent des sections en attente de développement fonctionnel ; leur gestion de profil est prête à recevoir les préférences d'interface qui seront ajoutées avec leurs futurs formulaires.
