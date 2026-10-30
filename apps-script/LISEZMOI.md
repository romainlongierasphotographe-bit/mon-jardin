# Installer le script des photos

Les photos de l'appli sont enregistrées dans ton Google Drive (dossier « Mon Jardin – Photos »).
L'onglet « Photos » du Google Sheet garde la liste : plante ou massif, date, lien.
Ce script est séparé du script existant de l'appli, qui reste tel quel.

1. Va sur https://script.google.com et clique sur « Nouveau projet ».
2. Renomme le projet « Mon Jardin – Photos » (en haut à gauche).
3. Efface le contenu du fichier `Code.gs`, puis colle tout le contenu de `photos.gs`. Enregistre (icône disquette).
4. Dans la liste des fonctions en haut, choisis `autoriser`, puis clique sur « Exécuter ».
   Google demande l'autorisation d'accéder au Drive et au Sheet. Accepte.
   Si un écran « Google n'a pas validé cette application » s'affiche, clique sur « Paramètres avancés », puis « Accéder à Mon Jardin – Photos ».
5. Clique sur « Déployer » → « Nouveau déploiement ».
   - Type : « Application Web » (roue dentée à côté de « Sélectionner le type »)
   - Exécuter en tant que : « Moi »
   - Qui a accès : « Tout le monde »
6. Clique sur « Déployer » et copie l'URL de l'application Web (elle se termine par `/exec`).
7. Colle cette URL dans `index.html`, dans la ligne `const PHOTO_SCRIPT_URL = '';`.

Si tu modifies le script plus tard : « Déployer » → « Gérer les déploiements » → crayon → Version : « Nouvelle version ». L'URL ne change pas.

## Remplissage automatique des fiches

Le même script remplit les fiches plantes à partir du nom de l'espèce, grâce à une IA.
Il faut une clé d'accès, à ajouter dans le script (elle n'apparaît jamais dans l'appli) :

- **Gemini (Google, offre gratuite)** : crée une clé sur https://aistudio.google.com/apikey
- **ou Claude (Anthropic, payant à l'usage)** : crée une clé sur https://console.anthropic.com

Dans l'éditeur Apps Script : roue dentée « Paramètres du projet » → « Propriétés du script » → « Ajouter une propriété » :
- Propriété : `GEMINI_API_KEY` (ou `ANTHROPIC_API_KEY`), Valeur : ta clé.

Ensuite, choisis la fonction `testerFiche` et clique sur « Exécuter » (Google redemande une autorisation pour les connexions externes : accepte). La fiche d'un rosier doit s'afficher dans le journal. Puis « Déployer » → « Gérer les déploiements » → crayon → « Nouvelle version » → « Déployer ».
