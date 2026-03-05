# Duplication d'arborescence (Google Drive Add-on)

Ce module complémentaire pour Google Drive permet de dupliquer des structures de dossiers complexes, avec une gestion spécifique pour les volumes importants.

## Français

### Description
Cet outil facilite la copie récursive de dossiers dans Google Drive. Il est particulièrement utile pour les arborescences profondes ou volumineuses grâce à son système de reprise après interruption.

### Fonctionnalités
* **Copie récursive** : Duplique l'intégralité de la structure des dossiers sélectionnés.
* **Option de copie des fichiers** : Possibilité d'inclure les fichiers ou de ne copier que la structure des dossiers.
* **Destinations flexibles** : Création de la copie au même emplacement, à la racine du Drive, ou dans un dossier spécifique via son identifiant (ID).
* **Filtres d'exclusion** : Permet d'ignorer certains sous-dossiers en saisissant leurs noms.
* **Gestion des limites de temps** : Si le traitement dépasse les limites de Google Apps Script, une option permet de reprendre la tâche là où elle s'est arrêtée.
* **Historique des opérations** : Affiche les 10 dernières duplications avec des liens directs vers les dossiers créés.

### Installation et configuration
* **Version** : v1.31.
* **Dépendances** : Nécessite l'activation du service "Drive API v3" dans le projet Google Apps Script.
* **Autorisation** : Une fonction `authorizeNow` est incluse pour valider les accès lors de la première utilisation.

### Licence
Ce projet est sous licence GNU GPL v3.

---

## English

### Description
This Google Drive add-on facilitates the recursive copying of folder structures. It is specifically designed to handle large or deep folder trees by providing a resume system to bypass execution time limits.

### Features
* **Recursive duplication**: Copies the entire structure of the selected folders.
* **File copy option**: Choose to include files or only replicate the folder structure.
* **Flexible destinations**: Create the copy in the same location, at the Drive root, or within a specific folder using its ID.
* **Exclusion filters**: Skip specific sub-folders by entering their names.
* **Time limit management**: If the process exceeds Google Apps Script limits, you can resume the task from the last point reached.
* **Operation history**: Keeps track of the last 10 duplications with direct links to the generated folders.

### Technical details
* **Version**: v1.31.
* **Dependencies**: Requires the "Drive API v3" advanced service to be enabled in the Apps Script project.
* **Authorization**: An `authorizeNow` function is provided for initial service authorization.

### License
This project is licensed under the GNU GPL v3.
