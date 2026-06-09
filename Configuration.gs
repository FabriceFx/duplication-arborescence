/**
 * ============================================================================
 * MODULE 1 : CONFIGURATION ET INTERNATIONALISATION (i18n)
 * ============================================================================
 */

/**
 * @typedef {Object} TaskStats
 * @property {number} dossiers - Nombre de dossiers créés.
 * @property {number} fichiers - Nombre de fichiers copiés.
 * @property {number} erreurs - Nombre d'erreurs rencontrées.
 * @property {number} ignores - Nombre d'éléments ignorés par les filtres.
 */

/**
 * @typedef {Object} DuplicationTask
 * @property {string} idDossierSource - ID du dossier source.
 * @property {string} nomSource - Nom du dossier source.
 * @property {string} idDossierDest - ID du dossier de destination.
 * @property {string} nomDestination - Nom personnalisé du dossier de destination.
 * @property {boolean} copierFichiers - Flag indiquant s'il faut copier les fichiers.
 * @property {boolean} executerArriereplan - Mode asynchrone par triggers temporels.
 * @property {boolean} conserverDroits - Flag pour la synchronisation ACL.
 * @property {string} emailUtilisateur - Email du propriétaire du job.
 * @property {string[]} exclusions - Liste des termes ou patterns à exclure.
 * @property {boolean} utiliserRegex - Si vrai, traite les exclusions comme des RegExp.
 * @property {TaskStats} statistiques - Compteurs d'exécution.
 * @property {Array<[string, string]>} fileAttente - Queue BFS contenant des couples [idSource, idDest].
 * @property {number} [indexFile] - Index de lecture optimisé pour le parcours de la file.
 */

const CONFIG = Object.freeze({
  VERSION: "v4.4 - Refactored June 2026",
  AUTEUR: "Fabrice FAUCHEUX (faucheux.bzh)",
  DUREE_MAX_MS_UI: 25 * 1000,
  DUREE_MAX_MS_ARRPLAN: 250 * 1000,
  TAILLE_SEGMENT_PROPRIETE: 8000,
  MAX_ERREURS_STOCKEES: 50,
  MAX_ENTREES_HISTORIQUE: 10,
  SEUIL_LIMITE_APERCU: 1000,
  MIME_DOSSIER: "application/vnd.google-apps.folder",
  MIME_RACCOURCI: "application/vnd.google-apps.shortcut",
  REGEX_VALIDATION_ID: /^[a-zA-Z0-9_-]+$/,
  COULEURS: Object.freeze({
    PRIMAIRE: "#1A73E8",
    SUR_SURFACE: "#202124",
    SUR_SURFACE_VARIANTE: "#5F6368",
    ERREUR: "#D93025",
    SUCCES: "#34A853",
    AVERTISSEMENT: "#EA8600",
    SURFACE_VARIANTE: "#F1F3F4"
  }),
  ICONES: Object.freeze({
    OUVRIR_NOUVEAU: "https://www.gstatic.com/images/icons/material/system/1x/open_in_new_googblue_18dp.png",
    LOGO: "https://www.gstatic.com/images/icons/material/system/2x/folder_copy_googblue_48dp.png"
  })
});

const DICTIONNAIRE_I18N = Object.freeze({
  TITLE: { fr: "Dupliquer une arborescence", en: "Duplicate a folder tree" },
  SUBTITLE: { fr: "Copie d'arborescences Google Drive", en: "Google Drive folder tree copy" },
  SELECT_FOLDER_PROMPT: { fr: "Sélectionnez un dossier dans Google Drive pour commencer.", en: "Select a folder in Google Drive to start." },
  WARNING: { fr: "Attention", en: "Warning" },
  NOT_A_FOLDER: { fr: "Veuillez sélectionner un <b>dossier</b>, pas un fichier.", en: "Please select a <b>folder</b>, not a file." },
  SOURCE_FOLDER: { fr: "Dossier source", en: "Source folder" },
  PREVIEW_BTN: { fr: "Aperçu", en: "Preview" },
  CONFIG_HEADER: { fr: "Configuration", en: "Settings" },
  COPY_NAME_LABEL: { fr: "Nom de la copie", en: "Copy name" },
  COPY_OF: { fr: "Copie de %s", en: "Copy of %s" },
  DESTINATION_TITLE: { fr: "Destination", en: "Destination" },
  DEST_SAME: { fr: "Même emplacement", en: "Same location" },
  DEST_ROOT: { fr: "Racine de Mon Drive", en: "My Drive root" },
  DEST_CUSTOM: { fr: "Autre dossier (coller son ID)", en: "Other folder (paste its ID)" },
  DEST_CUSTOM_INT: { fr: "ID du dossier de destination", en: "Destination folder ID" },
  DEST_CUSTOM_HINT: { fr: "Visible dans l'URL : drive.google.com/drive/folders/[ID]", en: "Visible in the URL: drive.google.com/drive/folders/[ID]" },
  COPY_FILES_OPT: { fr: "Copier les fichiers en plus des dossiers", en: "Copy files in addition to folders" },
  RUN_BG_OPT: { fr: "Exécuter en arrière-plan (recommandé)", en: "Run in background (recommended)" },
  ADVANCED_OPTS: { fr: "Options avancées", en: "Advanced options" },
  EXCLUSIONS_LABEL: { fr: "Dossiers à ignorer", en: "Folders to ignore" },
  EXCLUSIONS_HINT: { fr: "Noms séparés par des virgules", en: "Comma-separated names" },
  USE_REGEX_OPT: { fr: "Utiliser les expressions régulières", en: "Use regular expressions" },
  SYNC_PERM_OPT: { fr: "Conserver les droits de partage (Plus lent)", en: "Preserve sharing permissions (Slower)" },
  START_BTN: { fr: "Lancer la duplication", en: "Start duplication" },
  RESUME_BTN: { fr: "Reprendre la duplication en cours", en: "Resume current duplication" },
  INVALID_FOLDER_ID: { fr: "ID de dossier invalide.", en: "Invalid folder ID." },
  REASONABLE_SIZE: { fr: "Taille raisonnable — duplication en une seule passe.", en: "Reasonable size — duplication in a single pass." },
  LARGE_TREE: { fr: "Arborescence volumineuse", en: "Large folder tree" },
  MULTIPLE_PASSES: { fr: "La duplication se fera en plusieurs passes automatiques.", en: "Duplication will be done in multiple automatic passes." },
  LIMITED_PREVIEW: { fr: "Aperçu limité (>%s éléments). L'arborescence réelle est plus grande.", en: "Limited preview (>%s items). Actual folder tree is larger." },
  SUBFOLDERS: { fr: "Sous-dossiers", en: "Subfolders" },
  FILES: { fr: "Fichiers", en: "Files" },
  DEPTH: { fr: "Profondeur", en: "Depth" },
  LEVELS: { fr: "%s niveau(x)", en: "%s level(s)" },
  BACK_BTN: { fr: "← Retour", en: "← Back" },
  UNABLE_TO_READ_FOLDER: { fr: "Impossible de lire le dossier : %s", en: "Unable to read folder: %s" },
  EMPTY_COPY_NAME: { fr: "Le nom de la copie ne peut pas être vide.", en: "The copy name cannot be empty." },
  INVALID_SOURCE_ID: { fr: "L'ID du dossier source est invalide.", en: "The source folder ID is invalid." },
  INVALID_DEST_CHARS: { fr: "L'ID de destination contient des caractères invalides.", en: "The destination folder ID contains invalid characters." },
  INVALID_DEST_ID: { fr: "ID de destination invalide ou inaccessible.", en: "Invalid or inaccessible destination folder ID." },
  HISTORY_HEADER: { fr: "Historique", en: "History" },
  HISTORY_EMPTY: { fr: "Aucune duplication effectuée pour l'instant.", en: "No duplication performed yet." },
  CLEAR_HISTORY_BTN: { fr: "Effacer l'historique", en: "Clear history" },
  HISTORY_CLEARED: { fr: "Historique effacé.", en: "History cleared." },
  HISTORY_OPEN_ALT: { fr: "Ouvrir le dossier", en: "Open folder" },
  BG_PROCESSING_TITLE: { fr: "Traitement en cours", en: "Processing" },
  BG_PROCESSING_SUB: { fr: "Duplication automatisée en arrière-plan", en: "Automated duplication in the background" },
  BG_PROCESSING_MSG1: { fr: "Le script a pris le relais en arrière-plan.", en: "The script has taken over in the background." },
  BG_PROCESSING_MSG2: { fr: "Vous pouvez fermer cet outil. Vous recevrez un email dès que la copie sera terminée.", en: "You can close this tool. You will receive an email as soon as the copy is finished." },
  OPEN_FOLDER_BUILD: { fr: "Ouvrir le dossier (en construction)", en: "Open folder (under construction)" },
  PASS_INCOMPLETE_TITLE: { fr: "Passe incomplète", en: "Incomplete pass" },
  PASS_INCOMPLETE_SUB: { fr: "Traitement interrompu par le temps limite", en: "Processing interrupted by time limit" },
  FOLDERS_CREATED: { fr: "Dossiers créés", en: "Folders created" },
  FILES_COPIED: { fr: "Fichiers copiés", en: "Files copied" },
  REMAINING_TO_PROCESS: { fr: "Restant à traiter", en: "Remaining to process" },
  FOLDERS_IN_QUEUE: { fr: "%s dossier(s) en file d'attente", en: "%s folder(s) in queue" },
  LARGE_TREE_CONTINUE: { fr: "L'arborescence est volumineuse. Cliquez sur <b>Continuer</b> pour la passe suivante.", en: "The folder tree is large. Click <b>Continue</b> for the next pass." },
  CONTINUE_BTN: { fr: "Continuer la duplication", en: "Continue duplication" },
  FOLDER_CREATED: { fr: "Dossier créé", en: "Folder created" },
  SUBFOLDERS_CREATED: { fr: "Sous-dossiers créés", en: "Subfolders created" },
  FOLDERS_IGNORED: { fr: "Dossiers ignorés", en: "Folders ignored" },
  ERRORS_ENCOUNTERED: { fr: "%s erreur(s) rencontrée(s) :", en: "%s error(s) encountered:" },
  AND_OTHERS: { fr: "…et %s autre(s).", en: "…and %s other(s)." },
  OPEN_CREATED_FOLDER: { fr: "Ouvrir le dossier créé", en: "Open created folder" },
  DUPLICATION_FINISHED_TITLE: { fr: "Duplication terminée", en: "Duplication finished" },
  FOLDERS_CREATED_SUCCESS: { fr: "%s dossier(s) créé(s) avec succès", en: "%s folder(s) successfully created" },
  ERROR_TITLE: { fr: "Erreur", en: "Error" },
  PROBLEM_OCCURRED: { fr: "Un problème est survenu", en: "A problem occurred" },
  NO_DUPLICATION_CORRUPTED: { fr: "Aucune duplication en cours ou données corrompues.", en: "No duplicate in progress or data corrupted." },
  EMAIL_SUCCESS_SUBJECT: { fr: "Duplication terminée : %s", en: "Duplication finished: %s" },
  EMAIL_SUCCESS_HELLO: { fr: "Bonjour", en: "Hello" },
  EMAIL_SUCCESS_BODY_1: { fr: "La duplication de votre arborescence est terminée avec succès.", en: "The duplication of your folder tree completed successfully." },
  EMAIL_SUCCESS_BODY_2: { fr: "La duplication de votre arborescence Google Drive est terminée. Voici le résumé :", en: "The duplication of your Google Drive folder tree is complete. Here is the summary:" },
  EMAIL_SUCCESS_ERR: { fr: "%s erreur(s) rencontrée(s) (voir les logs pour les détails).", en: "%s error(s) encountered (check the logs for details)." },
  EMAIL_SUCCESS_ERR_HTML: { fr: "⚠ %s erreur(s) rencontrée(s)", en: "⚠ %s error(s) encountered" },
  EMAIL_SUCCESS_FOLDER: { fr: "Dossier : %s", en: "Folder: %s" },
  EMAIL_SUCCESS_ACCESS: { fr: "Accéder au dossier : %s", en: "Access folder: %s" },
  EMAIL_SUCCESS_OPEN_BTN: { fr: "Ouvrir le dossier", en: "Open folder" },
  EMAIL_SUCCESS_SUBTITLE: { fr: "Votre arborescence a été dupliquée avec succès", en: "Your folder tree has been duplicated successfully" },
  EMAIL_FAILED_SUBJECT: { fr: "Échec de la duplication de %s", en: "Duplication failed for %s" },
  EMAIL_FAILED_BODY_1: { fr: "La duplication en arrière-plan s'est arrêtée suite à une erreur :", en: "The background duplication stopped due to an error:" },
  EMAIL_FAILED_BODY_1_HTML: { fr: "La duplication en arrière-plan s'est arrêtée suite à une erreur inattendue.", en: "The background duplication stopped due to an unexpected error." },
  EMAIL_FAILED_BODY_2: { fr: "Vous pouvez reprendre manuellement via l'interface de l'add-on.", en: "You can resume manually via the add-on interface." },
  EMAIL_FAILED_TITLE: { fr: "Échec de la duplication", en: "Duplication failed" },
  EMAIL_FAILED_DETAIL: { fr: "Détail de l'erreur", en: "Error details" },
  EMAIL_FAILED_ACTION_TITLE: { fr: "Action recommandée", en: "Recommended action" },
  EMAIL_FAILED_ACTION_DESC: { fr: "Ouvrez Google Drive, sélectionnez le dossier source et utilisez le bouton <strong>« Reprendre la duplication en cours »</strong> dans le panneau latéral de l'add-on.", en: "Open Google Drive, select the source folder and use the <strong>\"Resume current duplication\"</strong> button in the add-on side panel." },
  EMAIL_FOOTER_DESC: { fr: "Add-on Google Drive — Duplication d'arborescence", en: "Google Drive Add-on — Folder Tree Duplication" }
});

/**
 * Système de traduction optimisé.
 * @param {string} cle La clé dans le dictionnaire.
 * @param {...string} args Remplacements dynamiques (%s).
 * @return {string} Le texte final traduit.
 */
function t(cle, ...args) {
  const locale = (Session.getActiveUserLocale() || "en").toLowerCase();
  const langue = locale.startsWith('fr') ? 'fr' : 'en';
  
  if (!DICTIONNAIRE_I18N[cle]) {
    console.warn(`Clé manquante : ${cle}`);
    return cle;
  }
  
  let texte = DICTIONNAIRE_I18N[cle][langue] || DICTIONNAIRE_I18N[cle]['en'];
  for (const arg of args) {
    texte = texte.replace('%s', arg);
  }
  return texte;
}

