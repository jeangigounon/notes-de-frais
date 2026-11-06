# Notes de frais

Une petite application web, installable sur iPhone comme une vraie app, pour ne plus oublier ses notes de frais ni perdre ses tickets.

1. **Photographiez** un ticket ou déposez une facture (photo ou PDF).
2. **L'IA lit le justificatif** : commerçant, date, montant TVAC, TVA, catégorie… Vous corrigez si besoin.
3. **Choisissez le client** à qui l'envoyer.
4. **Envoyez** : l'e-mail part avec les photos et le fichier Excel **exactement au format demandé par ce client**.

Le téléphone sert à photographier et envoyer en deux gestes ; le PC à configurer confortablement les clients et leurs modèles Excel.

## Fonctions

- **Capture** : appareil photo, galerie, **factures PDF** (aperçu de la 1re page), glisser-déposer et copier-coller (Ctrl+V) sur PC. Les photos sont compressées et conservées sur l'appareil, même hors ligne.
- **Reconnaissance** : Claude (vision, PDF compris) pour une lecture fiable, ou une lecture locale gratuite en secours (texte intégré des PDF, sinon OCR Tesseract).
- **Note en cours par client** : on ajoute les tickets au fil de l'eau (« Ajouter à la note », « Ticket suivant »), puis on envoie le tout en une fois, dans **un seul fichier Excel**.
- **Devises** : chaque ticket garde sa devise ; le taux de change de référence de la BCE au jour de la dépense est récupéré automatiquement (modifiable) et les montants sont convertis en euros.
- **Clients** : adresse(s) e-mail, objet et message types avec variables (`{nom}`, `{periode}`, `{total}`, `{nb}`, `{client}`, `{date}`), couleur, notes.
- **Excel par client**, au choix :
  - **Son modèle** : déposez le fichier Excel du client (.xlsx, ou .xls converti automatiquement). **Claude l'analyse** (structure, formules, colonnes par type de dépense et par devise, cellules de taux de change) et propose où placer chaque information ; vous vérifiez sur une grille et corrigez d'un clic. Une colonne peut ne recevoir que certaines catégories (repas, transport, logement, « divers »…) et/ou une devise (€, CHF, £…). Le fichier est rempli sans toucher à sa mise en forme : logos, couleurs, bordures, cellules fusionnées, formules et autres feuilles restent intacts. Si le tableau est trop court, des lignes sont ajoutées et les formules de total s'étendent : tout reste dans un seul fichier.
  - **Tableau standard** : un tableau propre, colonnes au choix, avec totaux.
  - **Pas d'Excel** : seulement les justificatifs.
- **Envoi** :
  - **Microsoft 365** : envoi direct depuis votre adresse professionnelle (configuration unique, voir plus bas) ;
  - **Menu Partager** du téléphone (Outlook, Mail…) avec toutes les pièces jointes ;
  - **Brouillon Outlook** (.eml) sur PC, qui s'ouvre prêt à partir ;
  - archive **.zip**.
- **Rappel** : l'accueil affiche ce qui attend d'être envoyé, par client, et signale les tickets oubliés depuis plus d'une semaine. Une pastille sur l'icône de l'app compte les justificatifs en attente.
- **Synchronisation PC ↔ téléphone** des clients (modèles Excel compris) via un Gist GitHub secret, et **QR code** pour configurer le téléphone sans rien retaper.
- **Sauvegarde / restauration** complète en un fichier.

## Mise en route

### 1. Installer sur l'iPhone

Ouvrez l'adresse du site dans **Safari** → bouton **Partager** → **Sur l'écran d'accueil**. L'app « Frais » s'ouvre en plein écran et conserve durablement vos tickets.

### 2. Clé Claude (reconnaissance des tickets)

1. Créez une clé sur [console.anthropic.com](https://console.anthropic.com/settings/keys) et ajoutez quelques euros de crédit.
2. Collez-la dans **Réglages → Reconnaissance des tickets**, puis « Tester la clé ».

Coût indicatif par ticket : environ 2 centimes avec Claude Opus 5.5 (par défaut, le plus précis), 1 centime avec Sonnet 5.5, moins de 0,2 centime avec Haiku 5.5. La clé n'est enregistrée que sur l'appareil.

### 3. Synchroniser le PC et le téléphone

1. [Créez un jeton GitHub](https://github.com/settings/tokens/new?scopes=gist&description=Notes%20de%20frais) avec uniquement l'autorisation **gist**.
2. Collez-le dans **Réglages → Synchronisation** sur le PC.
3. Sur le PC, **Réglages → Afficher le QR code de configuration**, puis scannez-le avec l'appareil photo de l'iPhone : clé Claude, synchronisation et profil sont copiés.

Les clients et le profil sont alors partagés entre vos appareils. Les photos des tickets restent sur l'appareil qui les a prises.

### 4. Envoi direct via Microsoft 365 (facultatif)

Sans cette étape, l'envoi passe par le menu Partager (téléphone) ou un brouillon Outlook (PC).

1. [entra.microsoft.com → Inscriptions d'applications](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **Nouvelle inscription**.
2. Nom « Notes de frais », comptes de **cet annuaire d'organisation uniquement**.
3. URI de redirection : plateforme **Application monopage (SPA)**, avec l'adresse exacte du site (affichée dans Réglages).
4. Copiez l'**ID d'application (client)** dans **Réglages → Microsoft 365**, vérifiez la permission déléguée **Microsoft Graph › Mail.Send**, puis « Se connecter ».

Si votre organisation exige l'accord d'un administrateur, transmettez ces informations à votre service informatique.

## Confidentialité

- Aucun serveur : tout fonctionne dans le navigateur. Les justificatifs sont stockés localement (IndexedDB).
- Les photos ne sont envoyées qu'à l'API Claude (pour la lecture) et au destinataire de l'e-mail.
- Les clés (Claude, GitHub, Microsoft) restent dans le stockage local de chaque appareil. Le QR code de configuration les contient : ne l'affichez pas devant d'autres personnes.

## Développement

Site statique sans étape de compilation (HTML, CSS et modules JavaScript). Bibliothèques chargées à la demande depuis un CDN, en versions figées : JSZip (Excel), SheetJS (conversion des .xls), PDF.js (factures PDF), Tesseract.js (lecture locale), qrcode-generator, SDK Anthropic. Taux de change : [frankfurter.dev](https://frankfurter.dev) (taux de référence BCE).

Pour un modèle client au format .xls, la conversion dans le navigateur garde formules, fusions et formats de nombres mais pas les couleurs ni les bordures : pour tout conserver, ouvrez-le dans Excel et enregistrez-le en .xlsx avant de le déposer.

```
python -m http.server 8000
```

puis ouvrir http://localhost:8000. Publication : GitHub Pages sur la branche `main`, dossier racine.

```
index.html, manifest.webmanifest, sw.js   coquille, PWA, hors ligne
css/app.css                               design
js/app.js                                 démarrage et navigation
js/store.js                               données (localStorage + IndexedDB)
js/ai.js                                  reconnaissance des tickets (Claude / lecture locale)
js/templateAI.js                          analyse des modèles Excel par Claude
js/xlsx.js                                remplissage des modèles, ajout de lignes, tableau standard, .xls
js/pdf.js, js/fx.js                       factures PDF, taux de change BCE
js/mail.js                                e-mail : Microsoft Graph, Partager, .eml, .zip
js/sync.js, js/transfer.js                synchronisation Gist, QR code, sauvegarde
js/views/                                 écrans
```
