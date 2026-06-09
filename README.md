# 📦 Duplication d'Arborescence — Google Drive Add-on (v4.4)

[🇫🇷 Version Française](#-version-française) | [🇬🇧 English Version](#-english-version)

---

## 🇫🇷 Version Française

> Un add-on professionnel pour Google Drive permettant de dupliquer l'arborescence complète d'un dossier (sous-dossiers et fichiers optionnels) vers un emplacement de destination, avec gestion automatique des timeouts pour les arborescences volumineuses.

<a href="https://developers.google.com/apps-script"><img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white" alt="Google Apps Script"></a>
<a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-indigo?style=for-the-badge" alt="License: MIT"></a>
<a href="README.md"><img src="https://img.shields.io/badge/Status-Production-brightgreen?style=for-the-badge" alt="Status: Production"></a>

---

### ✨ Fonctionnalités clés

- 📁 **Copie Intégrale de Structure** : Duplique récursivement tous les sous-dossiers et fichiers (optionnel) avec conservation de la hiérarchie.
- ⏱️ **Résistance au Timeout GAS (6 min)** : Découpe le travail en passes successives. Sauvegarde l'état actuel et programme un déclencheur temporel (`trigger`) pour reprendre automatiquement 1 min plus tard si nécessaire.
- 🔎 **Algorithme BFS (Breadth-First Search)** : Parcours prévisible et uniforme par niveau de profondeur.
- 🔏 **Filtres d'Exclusions** : Possibilité d'ignorer certains sous-dossiers par listes de noms ou via expressions régulières (Regex).
- 🔑 **Synchronisation des Droits** : Conserve optionnellement les partages et droits d'accès des dossiers d'origine.
- 📧 **Notification de Fin** : Envoie un email de récapitulatif HTML soigné dès que la duplication complète est achevée.
- 🌍 **Interface Bilingue (i18n)** : L'interface s'adapte automatiquement en Français ou en Anglais selon vos paramètres régionaux Google.

---

### 🚀 Installation & configuration

1. Ouvrez votre console Google Apps Script ou créez un projet d'Add-on Google Workspace.
2. Copiez le code source dans quatre fichiers distincts : `Configuration.gs`, `Persistance.gs`, `MoteurDuplication.gs` et `Interface.gs`.
3. Configurez votre fichier manifeste **[appsscript.json](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/appsscript.json)** pour inclure l'extension Google Drive et déclarer les scopes d'accès.
4. Déployez l'Add-on en mode test ou installez-le sur votre compte.
5. Sélectionnez n'importe quel dossier dans Google Drive : le panneau latéral **"Dupliquer une arborescence"** apparaît !

---

### 🛠️ Structure du Projet

L'application est structurée en 4 modules métiers pour des performances et une maintenabilité optimales :
- **[Configuration.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/Configuration.gs)** : Constantes, paramètres et dictionnaire de traductions (i18n).
- **[Persistance.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/Persistance.gs)** : Sauvegarde par segments via PropertiesService et gestion des déclencheurs d'arrière-plan.
- **[MoteurDuplication.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/MoteurDuplication.gs)** : L'algorithme BFS optimisé en O(N) pour la copie massive de l'arborescence.
- **[Interface.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/Interface.gs)** : Gestion de l'UI Google Workspace (CardService) et notifications email.
- **[appsscript.json](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/appsscript.json)** : Manifeste de l'extension spécifiant les dépendances d'API et les scopes OAuth Google Drive requis.

---

### 👤 Auteur

- **[Fabrice Faucheux](https://faucheux.bzh)** (FF Labs) - [GitHub](https://github.com/FabriceFx)

---

### 📄 Licence

Ce projet est disponible sous licence **MIT**.

---

---

## 🇬🇧 English Version

> A professional Google Drive Add-on to duplicate complete folder structures (with optional subfolders and files copying) into custom locations, featuring automatic time-limit crash recovery for very large directories.

<a href="https://developers.google.com/apps-script"><img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white" alt="Google Apps Script"></a>
<a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-indigo?style=for-the-badge" alt="License: MIT"></a>
<a href="README.md"><img src="https://img.shields.io/badge/Status-Production-brightgreen?style=for-the-badge" alt="Status: Production"></a>

---

### ✨ Key Features

- 📁 **Comprehensive Hierarchy Duplication**: Recursively duplicates all subfolders and files while fully preserving directory structures.
- ⏱️ **GAS 6-minute Timeout Shield**: Splits processing into chunks. Saves current state and schedules background triggers to automatically resume 1 minute later if execution limits are hit.
- 🔎 **BFS Traversal Engine**: Navigates folders level-by-level (Breadth-First Search) for reliable and uniform queue processing.
- 🔏 **Custom Exclusion Filters**: Skip specific folders by name lists or utilizing regular expressions (Regex).
- 🔑 **Permission Sync**: Optional deep copy of original sharing accesses and permissions for the newly generated directories.
- 📧 **Completion Emails**: Automatically dispatches a well-designed summary HTML email once the entire folder structure is successfully copied.
- 🌍 **Native Translation Dictionary (i18n)**: Sidebar components automatically switch between French and English based on the active Google account's locale settings.

---

### 🚀 Installation & Setup

1. Open your Google Apps Script editor or build a Google Workspace Add-on project.
2. Place the source code into four separate files: `Configuration.gs`, `Persistance.gs`, `MoteurDuplication.gs`, and `Interface.gs`.
3. Configure your **[appsscript.json](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/appsscript.json)** manifest file to support Google Drive integrations and request mandatory scopes.
4. Deploy the add-on for testing or install it globally on your workspace account.
5. Select any directory inside Google Drive: the custom **"Duplicate a folder tree"** side panel will launch!

---

### 🛠️ Project Structure

The application is structured into 4 distinct modules for maximum performance and maintainability:
- **[Configuration.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/Configuration.gs)**: Constants, settings, and native translation dictionary (i18n).
- **[Persistance.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/Persistance.gs)**: Chunk-based serialization via PropertiesService and background trigger management.
- **[MoteurDuplication.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/MoteurDuplication.gs)**: The highly optimized O(N) BFS algorithm for massive folder tree duplication.
- **[Interface.gs](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/Interface.gs)**: Google Workspace UI generation (CardService) and email notifications.
- **[appsscript.json](file:///Users/fabrice/Documents/Mes%20développements/Duplication%20d'arborescence/appsscript.json)**: The extension manifest specifying Google Drive add-on contexts and mandatory API OAuth scopes.

---

### 👤 Author

- **[Fabrice Faucheux](https://faucheux.bzh)** (FF Labs) - [GitHub](https://github.com/FabriceFx)

---

### 📄 License

This project is licensed under the terms of the **MIT License**.

---

---
<p align="center"><a href="https://faucheux.bzh" target="_blank" style="color: inherit; text-decoration: none;">&lt;&gt; par Fabrice Faucheux</a></p>
