# 📂 Dupliquer une arborescence | Duplicate Folder Tree
> Un Add-on Google Workspace puissant pour cloner des structures de dossiers complexes sans effort.

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white)](https://developers.google.com/apps-script)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg?style=for-the-badge)](https://www.gnu.org/licenses/gpl-3.0)
[![Version](https://img.shields.io/badge/version-4.3-green?style=for-the-badge)](https://github.com)

**Dupliquer une arborescence** est un module complémentaire pour Google Drive conçu pour surmonter les limitations natives de Google. Il permet de copier intégralement une structure de dossiers, avec ou sans fichiers, tout en gérant intelligemment les dépassements de temps (timeout) grâce à un moteur de traitement en arrière-plan.

## ✨ Points forts (Features)

- **🚀 Performance & robustesse** : Utilise un algorithme **BFS (Breadth-First Search)** pour un parcours stable des dossiers, évitant les erreurs de récursion.
- **🔄 Reprise automatique** : Système de *checkpoint* sauvegardant l'état du job dans le `PropertiesService` par segments (chunks) pour contourner la limite de 9 Ko.
- **⏱️ Gestion du timeout** : Détection automatique des limites d'exécution (6 min sur Google Apps Script) avec programmation de triggers de relance.
- **🛡️ Options avancées** :
  - Filtrage par **expressions régulières (Regex)** ou noms exacts.
  - Synchronisation des **droits de partage** (permissions).
  - Préservation des descriptions et des couleurs de dossiers.
- **📧 Notification Email** : Rapport détaillé envoyé automatiquement une fois la duplication terminée en arrière-plan.
- **📜 Historique** : Accès rapide aux 10 dernières opérations effectuées.

## 🛠️ Installation & prérequis

### Prérequis
- Un compte Google (Personnel ou Google Workspace).
- L'API Google Drive activée sur votre projet Apps Script.

### Étapes de déploiement
1. Ouvrez [Google Apps Script](https://script.google.com).
2. Créez un nouveau projet et collez le contenu de `Code.gs`.
3. Activez le service avancé : **Services > Drive API v3**.
4. Modifiez le fichier `appsscript.json` (Paramètres du projet > Afficher le fichier manifeste) avec le contenu fourni.
5. **Déploiement** : 
   - Cliquez sur `Déployer` > `Nouveau déploiement`.
   - Sélectionnez `Add-on Google Workspace`.
6. **Autorisation** : Exécutez la fonction `autoriserMaintenant()` dans l'éditeur pour valider les scopes OAuth.

## 🚀 Utilisation

1. Sélectionnez un dossier dans votre interface **Google Drive**.
2. Lancez l'Add-on depuis le panneau latéral droit.
3. Configurez vos options (Nom, Destination, Exclusions).
4. Cliquez sur **Lancer la duplication**.
   - *Note : Pour les dossiers volumineux, cochez "Exécuter en arrière-plan".*

## 🏗️ Technologies utilisées

- **Langage** : Google Apps Script (JavaScript V8)
- **API** : Google Drive API v3 (Advanced Service)
- **UI** : Card Service (Material Design)
- **Persistence** : PropertiesService (User Properties)

## 🤝 Contribution & licence

Les contributions sont les bienvenues ! N'hésitez pas à ouvrir une *Issue* ou une *Pull Request*.

Ce projet est sous licence **GNU GPL v3**. Voir le fichier [LICENSE](./LICENSE) pour plus de détails.

## 👤 Auteur

- **Fabrice FAUCHEUX** - *Développement initial & Architecture*

# 📂 Duplicate Folder Tree | Dupliquer une Arborescence
> A professional Google Workspace Add-on to clone complex folder structures effortlessly.

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white)](https://developers.google.com/apps-script)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg?style=for-the-badge)](https://www.gnu.org/licenses/gpl-3.0)
[![Version](https://img.shields.io/badge/version-4.3-green?style=for-the-badge)](https://github.com)

**Duplicate Folder Tree** is a Google Drive add-on designed to overcome native Google limitations. It allows for the full copy of folder structures, with or without files, while intelligently managing execution timeouts through a background processing engine.

## ✨ Features

- **🚀 Performance & Robustness**: Uses a **BFS (Breadth-First Search)** algorithm for stable folder traversal, avoiding recursion stack overflows.
- **🔄 Automatic Resume**: Checkpoint system saves job state in `PropertiesService` using chunks to bypass the 9 KB property limit.
- **⏱️ Timeout Management**: Automatically detects execution limits (6 min on Google Apps Script) and schedules triggers to resume.
- **🛡️ Advanced Options**:
  - Filtering via **Regular Expressions (Regex)** or exact names.
  - **Sharing Permissions** synchronization.
  - Preserves folder descriptions and custom colors.
- **📧 Email Notification**: Detailed report automatically sent once the background duplication is complete.
- **📜 History**: Quick access to the last 10 duplication operations.

## 🛠️ Installation & Prerequisites

### Prerequisites
- A Google Account (Personal or Google Workspace).
- Google Drive API enabled on your Apps Script project.

### Deployment Steps
1. Open [Google Apps Script](https://script.google.com).
2. Create a new project and paste the `Code.gs` content.
3. Enable the advanced service: **Services > Drive API v3**.
4. Update the `appsscript.json` file (Project Settings > Show manifest file) with the provided content.
5. **Deployment**: 
   - Click `Deploy` > `New deployment`.
   - Select `Google Workspace Add-on`.
6. **Authorization**: Run the `autoriserMaintenant()` function in the editor to grant OAuth scopes.

## 🚀 Usage

1. Select a folder in your **Google Drive** interface.
2. Launch the Add-on from the right side panel.
3. Configure your settings (Name, Destination, Exclusions).
4. Click **Start duplication**.
   - *Note: For large directories, check "Run in background".*

## 🏗️ Tech Stack

- **Language**: Google Apps Script (JavaScript V8)
- **API**: Google Drive API v3 (Advanced Service)
- **UI**: Card Service (Material Design)
- **Persistence**: PropertiesService (User Properties)

## 🤝 Contribution & License

Contributions are welcome! Feel free to open an *Issue* or a *Pull Request*.

This project is licensed under **GNU GPL v3**. See the [LICENSE](./LICENSE) file for details.

## 👤 Author

- **Fabrice FAUCHEUX** - *Initial Development & Architecture*
