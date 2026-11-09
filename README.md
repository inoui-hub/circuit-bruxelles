# Mon itinéraire à Bruxelles

Atelier de classe : chaque élève construit une journée idéale à Bruxelles (lieux, heures, commentaires, classement), l'envoie à son professeur et à un camarade, et reçoit en retour la feuille de route d'un camarade.

- **Élèves** : `index.html` — connexion à la classe, saisie de l'itinéraire, envoi au professeur, envoi à un camarade, feuille de route reçue.
- **Enseignant** : `prof.html` — liste des élèves connectés, superposition des itinéraires (une couleur par élève), filtre par classe, synthèse, annulation d'un envoi.
- **Serveur de classe** : `collecte-google.gs` — Google Apps Script + Google Sheet. Les étapes de mise en place sont en tête du fichier.

Sans serveur (`ENDPOINT` vide dans `index.html`), l'élève télécharge un fichier `.json` que l'enseignant importe dans `prof.html`.

Cartes : © OpenStreetMap contributors. Recherche de lieux : Nominatim. Carte interactive : Leaflet.
