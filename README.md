# Dupliquer une arborescence — Google Drive Add-on

Un Add-on Google Drive permettant de dupliquer l'arborescence complète d'un dossier, avec gestion intelligente des grandes structures via reprise automatique en arrière-plan.

---

## Fonctionnalités

- **Duplication d'arborescence** — copie récursive de tous les sous-dossiers avec préservation des couleurs et descriptions
- **Copie des fichiers** — optionnelle, activable à la demande
- **Destinations flexibles** — même emplacement, racine de Mon Drive, ou dossier personnalisé via ID
- **Exclusions** — filtrage par nom exact ou expression régulière
- **Aperçu préalable** — statistiques de l'arborescence avant lancement (dossiers, fichiers, profondeur)
- **Mode arrière-plan** — traitement automatique des grandes arborescences avec reprise sur timeout
- **Notification par email** — confirmation à la fin du traitement en arrière-plan
- **Historique** — suivi des 10 dernières duplications avec accès direct aux dossiers créés
- **Compatibilité Drives partagés** — support complet via `supportsAllDrives`

---

## Installation

### Prérequis

- Un compte Google
- Accès à [Google Apps Script](https://script.google.com)

### Étapes

1. Créez un nouveau projet sur [script.google.com](https://script.google.com)
2. Copiez le contenu de `Code.gs` dans l'éditeur
3. Remplacez le contenu de `appsscript.json` par le fichier fourni
   > Pour afficher `appsscript.json` : *Paramètres du projet → Afficher le fichier manifeste*
4. Activez l'API Drive avancée : *Services → Drive API v3*
5. Cliquez sur **Déployer → Nouveau déploiement**
   - Type : *Add-on Google Workspace*
   - Accès : *Moi-même* (test) ou *Tout le monde* (production)
6. Installez le déploiement depuis *drive.google.com → Extensions → Add-ons*
7. Lancez `authorizeNow()` une première fois depuis l'éditeur pour déclencher les autorisations OAuth

---

## Utilisation

1. Dans Google Drive, sélectionnez un dossier
2. Ouvrez le panneau latéral de l'Add-on
3. Configurez les options :

| Option | Description |
|---|---|
| **Nom de la copie** | Nom du dossier destination (pré-rempli) |
| **Destination** | Même emplacement, racine, ou ID personnalisé |
| **Copier les fichiers** | Inclut les fichiers en plus des dossiers |
| **Mode arrière-plan** | Recommandé pour les grandes arborescences |
| **Dossiers à ignorer** | Noms séparés par des virgules |
| **Expressions régulières** | Active le filtrage regex sur les exclusions |

4. Cliquez sur **Lancer la duplication**
5. En mode arrière-plan, un email vous est envoyé à la fin du traitement

---

## Architecture technique

### Gestion du timeout

Google Apps Script impose une limite d'exécution de 6 minutes. Le script utilise deux seuils de sécurité :

| Mode | Seuil | Déclenchement |
|---|---|---|
| Interactif | 25 secondes | L'utilisateur attend dans l'UI |
| Arrière-plan | 4 minutes | Trigger automatique |

Lorsque le seuil est atteint, l'état est sauvegardé et un nouveau trigger est programmé automatiquement pour reprendre 1 minute plus tard.

### Persistance de l'état

`PropertiesService` est limité à 9 ko par propriété. L'état du job est découpé en chunks de 8 000 caractères :

```
job_chunks_count  →  nombre de chunks
job_chunk_0       →  début du JSON
job_chunk_1       →  suite...
```

### Moteur de duplication

La traversée de l'arborescence utilise une **file d'attente BFS** (largeur d'abord) plutôt qu'une récursion, ce qui évite tout risque de stack overflow sur les arborescences profondes et offre une progression visible niveau par niveau.

### Structure du code

```
onDriveItemsSelected()      Point d'entrée principal (UI)
previewFolder()             Aperçu rapide de l'arborescence
startDuplication()          Initialisation et lancement
runDuplicationJob()         Moteur de duplication (BFS)
scheduleBackgroundJob()     Création du trigger de reprise
processBackgroundJob()      Exécution en arrière-plan
resumeDuplication()         Reprise manuelle depuis l'UI
saveJobState/loadJobState   Persistance par chunks
buildXxxCard()              Constructeurs de cartes UI
saveToHistory()             Historique des duplications
```

---

## Permissions requises

| Scope OAuth | Utilisation |
|---|---|
| `drive` | Lecture et écriture dans Google Drive |
| `drive.addons.metadata.readonly` | Métadonnées de l'Add-on |
| `script.scriptapp` | Création et gestion des triggers |
| `script.send_mail` | Envoi de l'email de confirmation |
| `userinfo.email` | Récupération de l'adresse email |

---

## Limitations connues

- L'aperçu est limité à 1 000 éléments pour rester dans les délais d'exécution
- L'historique conserve les 10 dernières duplications uniquement
- Les raccourcis Drive (`application/vnd.google-apps.shortcut`) sont ignorés volontairement
- En cas de crash inattendu en arrière-plan, une reprise manuelle reste possible via le panneau latéral

---

## Licence

MIT — libre d'utilisation, de modification et de distribution.

---

*v3.1 · Fabrice FAUCHEUX*
