# Mon itinéraire à Bruxelles

Atelier de classe : chaque élève construit une journée idéale à Bruxelles (lieux, heures, commentaires, classement) ; l'enseignant superpose tous les circuits sur une carte.

- **Élèves** : `index.html` — saisie uniquement, envoi à l'enseignant.
- **Enseignant** : `prof.html` — récupère les envois, une couleur par élève, filtre par classe, synthèse.
- **Collecte (facultatif)** : `collecte-google.gs` — script Google Apps Script qui reçoit les envois dans un Google Sheet. Les étapes de mise en place sont dans le fichier.

Sans le script de collecte, l'élève télécharge un fichier `.json` que l'enseignant importe dans `prof.html`.

Cartes : © OpenStreetMap contributors. Recherche de lieux : Nominatim. Carte interactive : Leaflet.
