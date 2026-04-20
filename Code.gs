// ============================================================
// DUPLIQUER UNE ARBORESCENCE — Google Drive Add-on
// ============================================================
// Auteur  : Fabrice FAUCHEUX
// Version : v4.1 (Bilingue FR/EN)
// ============================================================
//
// DESCRIPTION
// -----------
// Add-on Google Drive permettant de dupliquer l'arborescence
// complète d'un dossier (sous-dossiers et fichiers optionnels)
// vers un emplacement au choix, avec gestion des grandes
// arborescences via un système de reprise automatique.
//
// FONCTIONNEMENT GÉNÉRAL
// ----------------------
// 1. L'utilisateur sélectionne un dossier dans Google Drive.
// 2. Il configure la duplication via le panneau latéral :
//      - Nom de la copie
//      - Destination (même emplacement, racine, ou ID custom)
//      - Copie des fichiers en plus des dossiers (optionnel)
//      - Dossiers à ignorer (noms ou expressions régulières)
//      - Mode arrière-plan recommandé pour les grandes arbo.
// 3. La duplication parcourt l'arborescence via une file
//    d'attente BFS (Breadth-First Search) pour un traitement
//    prévisible et uniforme des niveaux de profondeur.
//
// GESTION DU TIMEOUT
// ------------------
// Google Apps Script impose une limite d'exécution de 6 min.
// Ce script utilise deux seuils de sécurité :
//   - 25 s  en mode interactif (l'utilisateur attend)
//   - ~4 min en mode arrière-plan (trigger automatique)
// Si la limite est atteinte, l'état de la file d'attente est
// sauvegardé et un nouveau trigger est programmé pour reprendre
// automatiquement 1 minute plus tard.
//
// PERSISTANCE DE L'ÉTAT
// ----------------------
// L'état du job en cours (file d'attente, statistiques,
// options) est sérialisé en JSON et stocké dans
// PropertiesService par chunks de 8 000 caractères pour
// contourner la limite de 9 ko par propriété.
//
// SYSTÈME D'ARRIÈRE-PLAN
// -----------------------
// Lorsque l'arborescence est volumineuse, le script crée un
// trigger time-based via ScriptApp qui relance
// traiterTacheArriereplan() toutes les ~1 minute jusqu'à
// complétion. Un email de confirmation est envoyé à
// l'utilisateur à la fin du traitement.
//
// STRUCTURE DU CODE
// -----------------
//   surSelectionElementsDrive()    Point d'entrée principal (UI)
//   apercuDossier()               Aperçu rapide de l'arborescence
//   lancerDuplication()            Initialisation et lancement
//   executerDuplication()          Moteur de duplication (BFS)
//   programmerArriereplan()        Création du trigger de reprise
//   traiterTacheArriereplan()      Exécution en arrière-plan
//   reprendreDuplication()         Reprise manuelle depuis l'UI
//   sauvegarderEtat/chargerEtat   Persistance par chunks
//   construireXxxCarte()          Constructeurs de cartes UI
//   sauvegarderHistorique()       Historique des duplications
//
// PERMISSIONS REQUISES
// --------------------
//   drive                          Lecture/écriture Drive
//   drive.addons.metadata.readonly Métadonnées Add-on
//   script.send_mail               Envoi d'email de fin
//   userinfo.email                 Récupération de l'email
//
// ============================================================

// ============================================================
// I18N - DICTIONNAIRE DE TRADUCTION
// ============================================================

const I18N = {
  TITLE: {
    fr: "Dupliquer une arborescence",
    en: "Duplicate a folder tree"
  },
  SUBTITLE: {
    fr: "Copie d'arborescences Google Drive",
    en: "Google Drive folder tree copy"
  },
  SELECT_FOLDER_PROMPT: {
    fr: "Sélectionnez un dossier dans Google Drive pour commencer.",
    en: "Select a folder in Google Drive to start."
  },
  WARNING: {
    fr: "Attention",
    en: "Warning"
  },
  NOT_A_FOLDER: {
    fr: "Veuillez sélectionner un <b>dossier</b>, pas un fichier.",
    en: "Please select a <b>folder</b>, not a file."
  },
  SOURCE_FOLDER: {
    fr: "Dossier source",
    en: "Source folder"
  },
  PREVIEW_BTN: {
    fr: "Aperçu",
    en: "Preview"
  },
  CONFIG_HEADER: {
    fr: "Configuration",
    en: "Settings"
  },
  COPY_NAME_LABEL: {
    fr: "Nom de la copie",
    en: "Copy name"
  },
  COPY_OF: {
    fr: "Copie de %s",
    en: "Copy of %s"
  },
  DESTINATION_TITLE: {
    fr: "Destination",
    en: "Destination"
  },
  DEST_SAME: {
    fr: "Même emplacement",
    en: "Same location"
  },
  DEST_ROOT: {
    fr: "Racine de Mon Drive",
    en: "My Drive root"
  },
  DEST_CUSTOM: {
    fr: "Autre dossier (coller son ID)",
    en: "Other folder (paste its ID)"
  },
  DEST_CUSTOM_INT: {
    fr: "ID du dossier de destination",
    en: "Destination folder ID"
  },
  DEST_CUSTOM_HINT: {
    fr: "Visible dans l'URL : drive.google.com/drive/folders/[ID]",
    en: "Visible in the URL: drive.google.com/drive/folders/[ID]"
  },
  COPY_FILES_OPT: {
    fr: "Copier les fichiers en plus des dossiers",
    en: "Copy files in addition to folders"
  },
  RUN_BG_OPT: {
    fr: "Exécuter en arrière-plan (recommandé)",
    en: "Run in background (recommended)"
  },
  ADVANCED_OPTS: {
    fr: "Options avancées",
    en: "Advanced options"
  },
  EXCLUSIONS_LABEL: {
    fr: "Dossiers à ignorer",
    en: "Folders to ignore"
  },
  EXCLUSIONS_HINT: {
    fr: "Noms séparés par des virgules",
    en: "Comma-separated names"
  },
  USE_REGEX_OPT: {
    fr: "Utiliser les expressions régulières",
    en: "Use regular expressions"
  },
  SYNC_PERM_OPT: {
    fr: "Conserver les droits de partage (Plus lent)",
    en: "Preserve sharing permissions (Slower)"
  },
  START_BTN: {
    fr: "Lancer la duplication",
    en: "Start duplication"
  },
  RESUME_BTN: {
    fr: "Reprendre la duplication en cours",
    en: "Resume current duplication"
  },
  INVALID_FOLDER_ID: {
    fr: "ID de dossier invalide.",
    en: "Invalid folder ID."
  },
  REASONABLE_SIZE: {
    fr: "Taille raisonnable — duplication en une seule passe.",
    en: "Reasonable size — duplication in a single pass."
  },
  LARGE_TREE: {
    fr: "Arborescence volumineuse",
    en: "Large folder tree"
  },
  MULTIPLE_PASSES: {
    fr: "La duplication se fera en plusieurs passes automatiques.",
    en: "Duplication will be done in multiple automatic passes."
  },
  LIMITED_PREVIEW: {
    fr: "Aperçu limité (>${LIMITE_APERCU} éléments). L'arborescence réelle est plus grande.",
    en: "Limited preview (>${LIMITE_APERCU} items). Actual folder tree is larger."
  },
  SUBFOLDERS: {
    fr: "Sous-dossiers",
    en: "Subfolders"
  },
  FILES: {
    fr: "Fichiers",
    en: "Files"
  },
  DEPTH: {
    fr: "Profondeur",
    en: "Depth"
  },
  LEVELS: {
    fr: "%s niveau(x)",
    en: "%s level(s)"
  },
  BACK_BTN: {
    fr: "← Retour",
    en: "← Back"
  },
  UNABLE_TO_READ_FOLDER: {
    fr: "Impossible de lire le dossier : %s",
    en: "Unable to read folder: %s"
  },
  EMPTY_COPY_NAME: {
    fr: "Le nom de la copie ne peut pas être vide.",
    en: "The copy name cannot be empty."
  },
  INVALID_SOURCE_ID: {
    fr: "L'ID du dossier source est invalide.",
    en: "The source folder ID is invalid."
  },
  INVALID_DEST_CHARS: {
    fr: "L'ID de destination contient des caractères invalides.",
    en: "The destination folder ID contains invalid characters."
  },
  INVALID_DEST_ID: {
    fr: "ID de destination invalide ou inaccessible.",
    en: "Invalid or inaccessible destination folder ID."
  },
  HISTORY_HEADER: {
    fr: "Historique",
    en: "History"
  },
  HISTORY_EMPTY: {
    fr: "Aucune duplication effectuée pour l'instant.",
    en: "No duplication performed yet."
  },
  CLEAR_HISTORY_BTN: {
    fr: "Effacer l'historique",
    en: "Clear history"
  },
  HISTORY_CLEARED: {
    fr: "Historique effacé.",
    en: "History cleared."
  },
  HISTORY_OPEN_ALT: {
    fr: "Ouvrir le dossier",
    en: "Open folder"
  },
  BG_PROCESSING_TITLE: {
    fr: "Traitement en cours",
    en: "Processing"
  },
  BG_PROCESSING_SUB: {
    fr: "Duplication automatisée en arrière-plan",
    en: "Automated duplication in the background"
  },
  BG_PROCESSING_MSG1: {
    fr: "Le script a pris le relais en arrière-plan.",
    en: "The script has taken over in the background."
  },
  BG_PROCESSING_MSG2: {
    fr: "Vous pouvez fermer cet outil. Vous recevrez un email dès que la copie sera terminée.",
    en: "You can close this tool. You will receive an email as soon as the copy is finished."
  },
  OPEN_FOLDER_BUILD: {
    fr: "Ouvrir le dossier (en construction)",
    en: "Open folder (under construction)"
  },
  PASS_INCOMPLETE_TITLE: {
    fr: "Passe incomplète",
    en: "Incomplete pass"
  },
  PASS_INCOMPLETE_SUB: {
    fr: "Traitement interrompu par le temps limite",
    en: "Processing interrupted by time limit"
  },
  FOLDERS_CREATED: {
    fr: "Dossiers créés",
    en: "Folders created"
  },
  FILES_COPIED: {
    fr: "Fichiers copiés",
    en: "Files copied"
  },
  REMAINING_TO_PROCESS: {
    fr: "Restant à traiter",
    en: "Remaining to process"
  },
  FOLDERS_IN_QUEUE: {
    fr: "%s dossier(s) en file d'attente",
    en: "%s folder(s) in queue"
  },
  LARGE_TREE_CONTINUE: {
    fr: "L'arborescence est volumineuse. Cliquez sur <b>Continuer</b> pour la passe suivante.",
    en: "The folder tree is large. Click <b>Continue</b> for the next pass."
  },
  CONTINUE_BTN: {
    fr: "Continuer la duplication",
    en: "Continue duplication"
  },
  FOLDER_CREATED: {
    fr: "Dossier créé",
    en: "Folder created"
  },
  SUBFOLDERS_CREATED: {
    fr: "Sous-dossiers créés",
    en: "Subfolders created"
  },
  FOLDERS_IGNORED: {
    fr: "Dossiers ignorés",
    en: "Folders ignored"
  },
  ERRORS_ENCOUNTERED: {
    fr: "%s erreur(s) rencontrée(s) :",
    en: "%s error(s) encountered:"
  },
  AND_OTHERS: {
    fr: "…et %s autre(s).",
    en: "…and %s other(s)."
  },
  OPEN_CREATED_FOLDER: {
    fr: "Ouvrir le dossier créé",
    en: "Open created folder"
  },
  DUPLICATION_FINISHED_TITLE: {
    fr: "Duplication terminée",
    en: "Duplication finished"
  },
  FOLDERS_CREATED_SUCCESS: {
    fr: "%s dossier(s) créé(s) avec succès",
    en: "%s folder(s) successfully created"
  },
  ERROR_TITLE: {
    fr: "Erreur",
    en: "Error"
  },
  PROBLEM_OCCURRED: {
    fr: "Un problème est survenu",
    en: "A problem occurred"
  },
  NO_DUPLICATION_CORRUPTED: {
    fr: "Aucune duplication en cours ou données corrompues.",
    en: "No duplicate in progress or data corrupted."
  },
  EMAIL_SUCCESS_SUBJECT: {
    fr: "Duplication terminée : %s",
    en: "Duplication finished: %s"
  },
  EMAIL_SUCCESS_HELLO: {
    fr: "Bonjour",
    en: "Hello"
  },
  EMAIL_SUCCESS_BODY_1: {
    fr: "La duplication de votre arborescence est terminée avec succès.",
    en: "The duplication of your folder tree completed successfully."
  },
  EMAIL_SUCCESS_BODY_2: {
    fr: "La duplication de votre arborescence Google Drive est terminée. <br>Voici le résumé :",
    en: "The duplication of your Google Drive folder tree is complete. <br>Here is the summary:"
  },
  EMAIL_SUCCESS_ERR: {
    fr: "%s erreur(s) rencontrée(s) (voir les logs pour les détails).",
    en: "%s error(s) encountered (check the logs for details)."
  },
  EMAIL_SUCCESS_ERR_HTML: {
    fr: "⚠ %s erreur(s) rencontrée(s)",
    en: "⚠ %s error(s) encountered"
  },
  EMAIL_SUCCESS_FOLDER: {
    fr: "Dossier : %s",
    en: "Folder: %s"
  },
  EMAIL_SUCCESS_ACCESS: {
    fr: "Accéder au dossier : %s",
    en: "Access folder: %s"
  },
  EMAIL_SUCCESS_OPEN_BTN: {
    fr: "Ouvrir le dossier",
    en: "Open folder"
  },
  EMAIL_SUCCESS_SUBTITLE: {
    fr: "Votre arborescence a été dupliquée avec succès",
    en: "Your folder tree has been duplicated successfully"
  },
  EMAIL_FAILED_SUBJECT: {
    fr: "Échec de la duplication de %s",
    en: "Duplication failed for %s"
  },
  EMAIL_FAILED_BODY_1: {
    fr: "La duplication en arrière-plan s'est arrêtée suite à une erreur :",
    en: "The background duplication stopped due to an error:"
  },
  EMAIL_FAILED_BODY_1_HTML: {
    fr: "La duplication en arrière-plan s'est arrêtée suite à une erreur inattendue.",
    en: "The background duplication stopped due to an unexpected error."
  },
  EMAIL_FAILED_BODY_2: {
    fr: "Vous pouvez reprendre manuellement via l'interface de l'add-on.",
    en: "You can resume manually via the add-on interface."
  },
  EMAIL_FAILED_TITLE: {
    fr: "Échec de la duplication",
    en: "Duplication failed"
  },
  EMAIL_FAILED_DETAIL: {
    fr: "Détail de l'erreur",
    en: "Error details"
  },
  EMAIL_FAILED_ACTION_TITLE: {
    fr: "Action recommandée",
    en: "Recommended action"
  },
  EMAIL_FAILED_ACTION_DESC: {
    fr: "Ouvrez Google Drive, sélectionnez le dossier source et utilisez le bouton <strong>« Reprendre la duplication en cours »</strong> dans le panneau latéral de l'add-on.",
    en: "Open Google Drive, select the source folder and use the <strong>\"Resume current duplication\"</strong> button in the add-on side panel."
  },
  EMAIL_FOOTER_DESC: {
    fr: "Add-on Google Drive — Duplication d'arborescence",
    en: "Google Drive Add-on — Folder Tree Duplication"
  }
};

/**
 * Fonction de traduction avec remplacement dynamique (ex: t('COPY_OF', 'Nom')).
 * @param {string} cle La clé dans l'objet I18N.
 * @param {...string} args Les paramètres à substituer.
 * @return {string} La chaîne traduite.
 */
function t(cle, ...args) {
  const langueSys = Session.getActiveUserLocale() || "en";
  // On utilise le français si la locaIe commence par 'fr', sinon anglais par défaut
  const langue = langueSys.toLowerCase().startsWith('fr') ? 'fr' : 'en';
  
  if (!I18N[cle]) {
    console.warn(`Clé de traduction manquante : ${cle}`);
    return cle;
  }
  
  let texte = I18N[cle][langue] || I18N[cle]['en']; // Fallback sécu sur 'en'
  
  // Remplacement des %s
  for (const arg of args) {
    texte = texte.replace('%s', arg);
  }
  
  return texte;
}

// ============================================================
// CONSTANTES ET CONFIGURATION
// ============================================================

/** @const {!GoogleAppsScript.Properties.Properties} Propriétés utilisateur */
const PROPRIETES = PropertiesService.getUserProperties();

/** @const {number} Durée max d'exécution en mode interactif (25 s) */
const DUREE_MAX_MS_UI = 25 * 1000;

/** @const {number} Durée max d'exécution en arrière-plan (~4 min 10 s) */
const DUREE_MAX_MS_ARRPLAN = 250 * 1000;

/** @const {string} Version de l'add-on */
const VERSION = "v4.1";

/** @const {string} Auteur de l'add-on */
const AUTEUR = "Fabrice FAUCHEUX";

/** @const {number} Taille maximale d'un chunk de propriété (en caractères) */
const TAILLE_SEGMENT = 8000;

/** @const {number} Nombre max d'erreurs conservées dans la tâche */
const MAX_ERREURS = 50;

/** @const {number} Nombre max d'entrées dans l'historique */
const MAX_HISTORIQUE = 10;

/** @const {number} Seuil d'éléments pour l'aperçu avant troncature */
const LIMITE_APERCU = 1000;

/** @const {string} MIME type des dossiers Google Drive */
const MIME_DOSSIER = "application/vnd.google-apps.folder";

/** @const {string} MIME type des raccourcis Google Drive */
const MIME_RACCOURCI = "application/vnd.google-apps.shortcut";

/** @const {RegExp} Validation d'un ID Google Drive */
const REGEX_ID_DRIVE = /^[a-zA-Z0-9_-]+$/;

/**
 * Palette de couleurs Material Design (Google).
 * Utilisée dans les balises <font> des widgets TextParagraph.
 * @enum {string}
 */
const COULEURS_MD = {
  PRIMAIRE:             "#1A73E8",
  SUR_SURFACE:          "#202124",
  SUR_SURFACE_VARIANTE: "#5F6368",
  ERREUR:               "#D93025",
  SUCCES:               "#34A853",
  AVERTISSEMENT:        "#EA8600",
  SURFACE_VARIANTE:     "#F1F3F4"
};

/**
 * URLs des icônes Material Design hébergées par Google.
 * @enum {string}
 */
const ICONES_MD = {
  OUVRIR_NOUVEAU: "https://www.gstatic.com/images/icons/material/system/1x/open_in_new_googblue_18dp.png",
  LOGO:           "https://www.gstatic.com/images/icons/material/system/2x/folder_copy_googblue_48dp.png"
};

// ============================================================
// POINT D'ENTRÉE
// ============================================================

/**
 * Point d'entrée principal de l'add-on.
 * Appelé lorsqu'un élément est sélectionné dans Google Drive
 * ou lors de l'ouverture du panneau latéral.
 *
 * @param {!Object} e Événement contextuel Drive
 * @return {!Card} Carte principale de l'interface
 */
function onDriveItemsSelected(e) {
  const elements = e && e.drive && e.drive.selectedItems;

  const carte = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(t('TITLE'))
      .setImageUrl(ICONES_MD.LOGO)
      .setImageStyle(CardService.ImageStyle.CIRCLE));

  // Aucun élément sélectionné
  if (!elements || elements.length === 0) {
    carte.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(
          `<font color="${COULEURS_MD.SUR_SURFACE_VARIANTE}">` +
          t('SELECT_FOLDER_PROMPT') +
          "</font>"
        )));
    carte.addSection(construirePiedDePage());
    return carte.build();
  }

  const element = elements[0];

  // L'élément sélectionné n'est pas un dossier
  if (element.mimeType !== MIME_DOSSIER) {
    carte.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(
          `<font color="${COULEURS_MD.AVERTISSEMENT}"><b>${t('WARNING')}</b></font><br>` +
          t('NOT_A_FOLDER')
        )));
    carte.addSection(construirePiedDePage());
    return carte.build();
  }

  // Section : Dossier source
  carte.addSection(CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel(t('SOURCE_FOLDER'))
      .setText(element.title)
      .setWrapText(true)
      .setButton(CardService.newTextButton()
        .setText(t('PREVIEW_BTN'))
        .setOnClickAction(CardService.newAction()
          .setFunctionName("apercuDossier")
          .setParameters({ idDossier: element.id, nomDossier: element.title })))));

  // Section : Configuration principale
  carte.addSection(CardService.newCardSection()
    .setHeader(t('CONFIG_HEADER'))
    .addWidget(CardService.newTextInput()
      .setFieldName("nom_copie")
      .setTitle(t('COPY_NAME_LABEL'))
      .setValue(t('COPY_OF', element.title)))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setTitle(t('DESTINATION_TITLE'))
      .setFieldName("emplacement_dest")
      .addItem(t('DEST_SAME'), "meme", true)
      .addItem(t('DEST_ROOT'), "racine", false)
      .addItem(t('DEST_CUSTOM'), "personnalise", false))
    .addWidget(CardService.newTextInput()
      .setFieldName("id_dest_personnalise")
      .setTitle(t('DEST_CUSTOM_INT'))
      .setHint(t('DEST_CUSTOM_HINT')))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("copier_fichiers")
      .addItem(t('COPY_FILES_OPT'), "oui", false))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("arriere_plan")
      .addItem(t('RUN_BG_OPT'), "oui", true))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("conserver_droits")
      .addItem(t('SYNC_PERM_OPT'), "oui", false)));

  // Section : Options avancées (collapsible — Material Design progressive disclosure)
  carte.addSection(CardService.newCardSection()
    .setHeader(t('ADVANCED_OPTS'))
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0)
    .addWidget(CardService.newTextInput()
      .setFieldName("exclusions")
      .setTitle(t('EXCLUSIONS_LABEL'))
      .setHint(t('EXCLUSIONS_HINT')))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("utiliser_regex")
      .addItem(t('USE_REGEX_OPT'), "oui", false)));

  // Section : Actions
  const sectionActions = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText(t('START_BTN'))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("lancerDuplication")
        .setParameters({ idDossier: element.id, nomDossier: element.title })));

  if (chargerEtat()) {
    sectionActions.addWidget(CardService.newDivider());
    sectionActions.addWidget(CardService.newTextButton()
      .setText(t('RESUME_BTN'))
      .setOnClickAction(CardService.newAction()
        .setFunctionName("reprendreDuplication")));
  }

  carte.addSection(sectionActions);
  carte.addSection(construireSectionHistorique());
  carte.addSection(construirePiedDePage());

  return carte.build();
}

// ============================================================
// APERÇU LIMITÉ
// ============================================================

/**
 * Affiche un aperçu de l'arborescence sélectionnée.
 * Parcourt récursivement le contenu jusqu'à un seuil de
 * LIMITE_APERCU éléments.
 *
 * @param {!Object} e Événement d'action avec paramètres idDossier et nomDossier
 * @return {!ActionResponse} Réponse de navigation avec la carte d'aperçu
 */
function apercuDossier(e) {
  const { idDossier, nomDossier } = e.parameters;

  if (!estIdDriveValide_(idDossier)) {
    return construireNavigationPush_(construireCarteErreur(t('INVALID_FOLDER_ID')));
  }

  try {
    const statistiques = { dossiers: 0, fichiers: 0, profondeur: 0, limiteAtteinte: false };
    compterContenuDossier_(idDossier, statistiques, 0);

    const estVolumineux = statistiques.dossiers > 50 || statistiques.fichiers > 200 || statistiques.limiteAtteinte;
    const suffixe = statistiques.limiteAtteinte ? "+" : "";

    let analyse = `<font color="${COULEURS_MD.SUCCES}">` +
      t('REASONABLE_SIZE') + "</font>";

    if (estVolumineux) {
      analyse = `<font color="${COULEURS_MD.AVERTISSEMENT}"><b>${t('LARGE_TREE')}</b></font><br>` +
        t('MULTIPLE_PASSES');
    }

    if (statistiques.limiteAtteinte) {
      let msgLimite = t('LIMITED_PREVIEW');
      msgLimite = msgLimite.replace('>${LIMITE_APERCU}', `>${LIMITE_APERCU}`); // Fix string substitution 
      analyse += `<br><br><font color="${COULEURS_MD.SUR_SURFACE_VARIANTE}"><i>` +
        msgLimite + "</i></font>";
    }

    const carteApercu = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle(t('PREVIEW_BTN'))
        .setSubtitle(nomDossier)
        .setImageUrl(ICONES_MD.LOGO)
        .setImageStyle(CardService.ImageStyle.CIRCLE))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newDecoratedText()
          .setTopLabel(t('SUBFOLDERS')).setText(`${statistiques.dossiers}${suffixe}`))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel(t('FILES')).setText(`${statistiques.fichiers}${suffixe}`))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel(t('DEPTH')).setText(t('LEVELS', statistiques.profondeur)))
        .addWidget(CardService.newDivider())
        .addWidget(CardService.newTextParagraph().setText(analyse))
        .addWidget(CardService.newTextButton()
          .setText(t('BACK_BTN'))
          .setOnClickAction(CardService.newAction()
            .setFunctionName("retourArriere"))))
      .addSection(construirePiedDePage())
      .build();

    return construireNavigationPush_(carteApercu);

  } catch (err) {
    return construireNavigationPush_(
      construireCarteErreur(t('UNABLE_TO_READ_FOLDER', err.message))
    );
  }
}

/**
 * Retourne à la carte précédente dans la pile de navigation.
 *
 * @return {!ActionResponse} Réponse de navigation popCard
 */
function retourArriere() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popCard())
    .build();
}

/**
 * Encapsule une carte dans une ActionResponse pushCard.
 *
 * @param {!Card} carte Carte à empiler
 * @return {!ActionResponse} Réponse de navigation pushCard
 * @private
 */
function construireNavigationPush_(carte) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(carte))
    .build();
}

/**
 * Compte récursivement les dossiers et fichiers d'une arborescence.
 * S'arrête dès que le seuil LIMITE_APERCU est dépassé.
 *
 * @note Utilise la récursion car le parcours est borné par LIMITE_APERCU.
 *       Le risque de stack overflow est nul avec cette limite.
 *
 * @param {string} idDossier ID du dossier à analyser
 * @param {!Object} statistiques Objet de statistiques (muté in place)
 * @param {number} profondeur Profondeur actuelle dans l'arborescence
 * @private
 */
function compterContenuDossier_(idDossier, statistiques, profondeur) {
  if (statistiques.dossiers + statistiques.fichiers > LIMITE_APERCU || statistiques.limiteAtteinte) {
    statistiques.limiteAtteinte = true;
    return;
  }
  if (profondeur > statistiques.profondeur) {
    statistiques.profondeur = profondeur;
  }

  let jetonPage = null;
  const sousDossiers = [];

  do {
    const reponse = avecRetentative_(() => Drive.Files.list({
      q: `'${idDossier}' in parents and trashed = false`,
      fields: "nextPageToken, files(id, mimeType)",
      pageSize: 1000,
      pageToken: jetonPage,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    }));

    const elements = reponse.files || [];
    for (const element of elements) {
      if (element.mimeType === MIME_DOSSIER) {
        statistiques.dossiers++;
        sousDossiers.push(element.id);
      } else if (element.mimeType !== MIME_RACCOURCI) {
        statistiques.fichiers++;
      }
    }
    if (statistiques.dossiers + statistiques.fichiers > LIMITE_APERCU) {
      statistiques.limiteAtteinte = true;
    }
    jetonPage = reponse.nextPageToken;
  } while (jetonPage && !statistiques.limiteAtteinte);

  for (const idSousDossier of sousDossiers) {
    if (statistiques.limiteAtteinte) break;
    compterContenuDossier_(idSousDossier, statistiques, profondeur + 1);
  }
}

// ============================================================
// PERSISTANCE (CHUNKING) ET UTILITAIRES
// ============================================================

/**
 * Sauvegarde l'état de la tâche dans PropertiesService par chunks.
 * Contourne la limite de 9 Ko par propriété en découpant
 * le JSON en segments de TAILLE_SEGMENT caractères.
 *
 * @param {!Object} tache État complet de la tâche de duplication
 */
function sauvegarderEtat(tache) {
  supprimerEtat();
  const json = JSON.stringify(tache);
  const nbSegments = Math.ceil(json.length / TAILLE_SEGMENT);
  PROPRIETES.setProperty("nombre_segments_tache", nbSegments.toString());

  for (let i = 0; i < nbSegments; i++) {
    PROPRIETES.setProperty(
      `segment_tache_${i}`,
      json.substring(i * TAILLE_SEGMENT, (i + 1) * TAILLE_SEGMENT)
    );
  }
}

/**
 * Charge l'état de la tâche depuis PropertiesService.
 *
 * @return {?Object} La tâche restaurée, ou null si aucun état trouvé/corrompu
 */
function chargerEtat() {
  const chaineCompteur = PROPRIETES.getProperty("nombre_segments_tache");
  if (!chaineCompteur) return null;
  const compteur = parseInt(chaineCompteur, 10);

  let json = "";
  for (let i = 0; i < compteur; i++) {
    json += PROPRIETES.getProperty(`segment_tache_${i}`) || "";
  }
  try {
    return JSON.parse(json);
  } catch (e) {
    console.error("État de la tâche corrompu, suppression.", e);
    supprimerEtat();
    return null;
  }
}

/**
 * Supprime toutes les propriétés associées à la tâche en cours.
 */
function supprimerEtat() {
  const chaineCompteur = PROPRIETES.getProperty("nombre_segments_tache");
  if (chaineCompteur) {
    const compteur = parseInt(chaineCompteur, 10);
    for (let i = 0; i < compteur; i++) {
      PROPRIETES.deleteProperty(`segment_tache_${i}`);
    }
    PROPRIETES.deleteProperty("nombre_segments_tache");
  }
}

/**
 * Exécute une action avec retry exponentiel (backoff + jitter).
 *
 * @param {function(): T} action Fonction à exécuter
 * @param {number} [maxTentatives=3] Nombre maximum de tentatives
 * @return {T} Résultat de l'action
 * @throws {Error} Si toutes les tentatives échouent
 * @template T
 * @private
 */
function avecRetentative_(action, maxTentatives = 3) {
  for (let tentative = 1; tentative <= maxTentatives; tentative++) {
    try {
      return action();
    } catch (e) {
      if (tentative >= maxTentatives) {
        console.error(`Échec après ${maxTentatives} tentatives : ${e.message}`);
        throw e;
      }
      const delaiMs = Math.pow(2, tentative) * 1000 + Math.round(Math.random() * 500);
      console.warn(`Tentative ${tentative}/${maxTentatives} échouée. Retry dans ${delaiMs}ms...`);
      Utilities.sleep(delaiMs);
    }
  }
}

/**
 * Valide qu'une chaîne est un ID Google Drive valide.
 *
 * @param {string} id ID à valider
 * @return {boolean} true si l'ID est syntaxiquement correct
 * @private
 */
function estIdDriveValide_(id) {
  return typeof id === "string" && id.length > 0 && REGEX_ID_DRIVE.test(id);
}

/**
 * Extrait une valeur depuis entreesFormulaire avec une valeur par défaut.
 *
 * @param {!Object} entreesFormulaire Objet formInputs de l'événement
 * @param {string} nomChamp Nom du champ
 * @param {string} [valeurParDefaut=""] Valeur par défaut
 * @return {string} La valeur du champ ou la valeur par défaut
 * @private
 */
function obtenirValeurFormulaire_(entreesFormulaire, nomChamp, valeurParDefaut = "") {
  return (entreesFormulaire[nomChamp] && entreesFormulaire[nomChamp][0])
    ? entreesFormulaire[nomChamp][0]
    : valeurParDefaut;
}

/**
 * Vérifie si une checkbox est cochée dans entreesFormulaire.
 *
 * @param {!Object} entreesFormulaire Objet formInputs de l'événement
 * @param {string} nomChamp Nom du champ
 * @return {boolean} true si la checkbox est cochée
 * @private
 */
function estCoche_(entreesFormulaire, nomChamp) {
  return !!(entreesFormulaire[nomChamp] && entreesFormulaire[nomChamp][0] === "oui");
}

// ============================================================
// LANCEMENT ET DUPLICATION
// ============================================================

/**
 * Initialise et lance une duplication d'arborescence.
 * Crée le dossier destination, prépare la tâche et lance le moteur BFS.
 *
 * @param {!Object} e Événement d'action avec paramètres et formInputs
 * @return {!Card} Carte de résultat, progression ou erreur
 */
function lancerDuplication(e) {
  supprimerEtat();

  const { idDossier, nomDossier } = e.parameters;
  const entreesFormulaire = e.formInputs || {};

  // Extraction et validation des entrées
  const nomDestination = obtenirValeurFormulaire_(entreesFormulaire, "nom_copie", t('COPY_OF', nomDossier)).trim();
  const copierFichiers = estCoche_(entreesFormulaire, "copier_fichiers");
  const executerArriereplan = estCoche_(entreesFormulaire, "arriere_plan");
  const conserverDroits = estCoche_(entreesFormulaire, "conserver_droits");
  const emplacementDest = obtenirValeurFormulaire_(entreesFormulaire, "emplacement_dest", "meme");
  const idDestPersonnalise = obtenirValeurFormulaire_(entreesFormulaire, "id_dest_personnalise").trim();
  const exclusionsBrutes = obtenirValeurFormulaire_(entreesFormulaire, "exclusions");
  const utiliserRegex = estCoche_(entreesFormulaire, "utiliser_regex");

  const exclusions = exclusionsBrutes.split(",").map(s => s.trim()).filter(Boolean);

  // Validations
  if (!nomDestination) {
    return construireCarteErreur(t('EMPTY_COPY_NAME'));
  }
  if (!estIdDriveValide_(idDossier)) {
    return construireCarteErreur(t('INVALID_SOURCE_ID'));
  }

  try {
    const dossierSource = DriveApp.getFolderById(idDossier);
    let idDossierParent;

    if (emplacementDest === "personnalise" && idDestPersonnalise) {
      if (!estIdDriveValide_(idDestPersonnalise)) {
        return construireCarteErreur(t('INVALID_DEST_CHARS'));
      }
      try {
        DriveApp.getFolderById(idDestPersonnalise);
        idDossierParent = idDestPersonnalise;
      } catch (err) {
        return construireCarteErreur(t('INVALID_DEST_ID'));
      }
    } else if (emplacementDest === "racine") {
      idDossierParent = DriveApp.getRootFolder().getId();
    } else {
      const parents = dossierSource.getParents();
      idDossierParent = parents.hasNext()
        ? parents.next().getId()
        : DriveApp.getRootFolder().getId();
    }

    const dossierDestination = avecRetentative_(() => Drive.Files.create({
      name: nomDestination,
      mimeType: MIME_DOSSIER,
      parents: [idDossierParent]
    }, null, { supportsAllDrives: true }));

    if (conserverDroits) {
      synchroniserDroits_(idDossier, dossierDestination.id);
    }

    /** @type {{ dossiers: number, fichiers: number, erreurs: !Array<string>, ignores: number }} */
    const statistiques = { dossiers: 0, fichiers: 0, erreurs: [], ignores: 0 };

    const tache = {
      idDossierSource: idDossier,
      nomSource: nomDossier,
      idDossierDest: dossierDestination.id,
      nomDestination,
      copierFichiers,
      executerArriereplan,
      conserverDroits,
      emailUtilisateur: Session.getEffectiveUser().getEmail(),
      exclusions,
      utiliserRegex,
      statistiques,
      fileAttente: [[idDossier, dossierDestination.id]]
    };

    const resultat = executerDuplication(tache, new Date().getTime());

    if (resultat.termine) {
      supprimerEtat();
      sauvegarderHistorique(nomDossier, nomDestination, dossierDestination.id, resultat.statistiques);
      // Envoi systématique d'un email de confirmation
      envoyerEmailFin_(
        Session.getEffectiveUser().getEmail(),
        nomDestination, dossierDestination.id, resultat.statistiques
      );
      return construireCarteResultat(nomDestination, dossierDestination.id, resultat.statistiques, copierFichiers);
    } else {
      sauvegarderEtat(resultat.tache);
      if (executerArriereplan) {
        programmerArriereplan();
        return construireCarteArriereplan(dossierDestination.id);
      } else {
        return construireCarteProgression(resultat.tache);
      }
    }

  } catch (err) {
    supprimerEtat();
    return construireCarteErreur(err.message);
  }
}

/**
 * Moteur de duplication utilisant un parcours BFS (Breadth-First Search).
 * Traite la file d'attente de paires [idSource, idDest] en créant
 * les dossiers et en copiant les fichiers dans les destinations.
 *
 * @param {!Object} tache État de la tâche (muté en place)
 * @param {number} heureDebut Timestamp de début en millisecondes
 * @param {boolean} [estArriereplan=false] true si exécuté en arrière-plan
 * @return {{ termine: boolean, tache: !Object, statistiques: ?Object }}
 */
function executerDuplication(tache, heureDebut, estArriereplan = false) {
  const fileAttente = tache.fileAttente;
  const exclusionsCompilees = [];
  const dureeMaxExecution = estArriereplan ? DUREE_MAX_MS_ARRPLAN : DUREE_MAX_MS_UI;

  console.log(
    "Début passe duplication. File d'attente : %s dossier(s). Mode BG : %s",
    fileAttente.length, estArriereplan
  );

  // Compilation des exclusions
  if (tache.exclusions && tache.exclusions.length > 0) {
    for (const exclusion of tache.exclusions) {
      if (tache.utiliserRegex) {
        try {
          exclusionsCompilees.push(new RegExp(exclusion, "i"));
        } catch (e) {
          console.warn("Expression régulière invalide ignorée : " + exclusion);
        }
      } else {
        exclusionsCompilees.push(exclusion.toLowerCase());
      }
    }
  }

  /**
   * Ajoute une erreur aux statistiques (tronquée à 100 caractères max).
   * @param {string} msg Message d'erreur
   */
  const ajouterErreur = (msg) => {
    console.error("Erreur duplication : " + msg);
    if (tache.statistiques.erreurs.length < MAX_ERREURS) {
      tache.statistiques.erreurs.push(msg.length > 100 ? `${msg.substring(0, 97)}...` : msg);
    }
  };

  /**
   * Vérifie si un nom de dossier est dans la liste d'exclusion.
   * @param {string} nom Nom du dossier
   * @return {boolean}
   */
  const estExclu = (nom) => {
    if (exclusionsCompilees.length === 0) return false;
    for (const compilee of exclusionsCompilees) {
      if (tache.utiliserRegex && compilee instanceof RegExp) {
        if (compilee.test(nom)) return true;
      } else if (typeof compilee === "string") {
        if (nom.toLowerCase().trim() === compilee) return true;
      }
    }
    return false;
  };

  // Boucle BFS principale
  while (fileAttente.length > 0) {
    if (new Date().getTime() - heureDebut > dureeMaxExecution) {
      tache.fileAttente = fileAttente;
      console.log(
        "Temps limite atteint (%sms). Pause. Reste : %s dossier(s).",
        dureeMaxExecution, fileAttente.length
      );
      return { termine: false, tache };
    }

    // BFS : shift() pour traiter en largeur d'abord (et non pop/DFS)
    const [idSourceActuel, idDestActuel] = fileAttente.shift();
    let jetonPage = null;

    do {
      try {
        const reponse = avecRetentative_(() => Drive.Files.list({
          q: `'${idSourceActuel}' in parents and trashed = false`,
          fields: "nextPageToken, files(id, name, mimeType, description, folderColorRgb)",
          pageSize: 1000,
          pageToken: jetonPage,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true
        }));

        const enfants = reponse.files || [];

        for (const enfant of enfants) {
          if (enfant.mimeType === MIME_DOSSIER) {
            if (estExclu(enfant.name)) {
              tache.statistiques.ignores++;
              continue;
            }
            try {
              const metadonneesDossier = {
                name: enfant.name,
                mimeType: MIME_DOSSIER,
                parents: [idDestActuel]
              };
              if (enfant.description) metadonneesDossier.description = enfant.description;
              if (enfant.folderColorRgb) metadonneesDossier.folderColorRgb = enfant.folderColorRgb;

              const nouveauDossier = avecRetentative_(() =>
                Drive.Files.create(metadonneesDossier, null, { supportsAllDrives: true })
              );

              if (tache.conserverDroits) {
                synchroniserDroits_(enfant.id, nouveauDossier.id);
              }

              tache.statistiques.dossiers++;
              fileAttente.push([enfant.id, nouveauDossier.id]);
            } catch (errDossier) {
              ajouterErreur(`Création dossier "${enfant.name}" : ${errDossier.message}`);
            }

          } else if (enfant.mimeType === MIME_RACCOURCI) {
            // Les raccourcis Drive sont ignorés
            continue;
          } else {
            if (tache.copierFichiers) {
              try {
                const metadonneesFichier = { parents: [idDestActuel], name: enfant.name };
                if (enfant.description) metadonneesFichier.description = enfant.description;

                const nouveauFichier = avecRetentative_(() =>
                  Drive.Files.copy(metadonneesFichier, enfant.id, { supportsAllDrives: true })
                );

                if (tache.conserverDroits) {
                  synchroniserDroits_(enfant.id, nouveauFichier.id);
                }

                tache.statistiques.fichiers++;
              } catch (errFichier) {
                ajouterErreur(`Copie fichier "${enfant.name}" : ${errFichier.message}`);
              }
            }
          }
        }
        jetonPage = reponse.nextPageToken;
      } catch (e) {
        ajouterErreur("Lecture du dossier source échouée : " + e.message);
        break;
      }
    } while (jetonPage);
  }

  console.log("File d'attente vide. Duplication terminée.");
  return { termine: true, statistiques: tache.statistiques };
}

// ============================================================
// SYSTÈME D'ARRIÈRE-PLAN (TRIGGERS)
// ============================================================

/**
 * Programme un trigger time-based pour reprendre la duplication
 * automatiquement dans 1 minute. Supprime les anciens triggers
 * traiterTacheArriereplan existants.
 */
function programmerArriereplan() {
  console.log("Programmation d'un déclencheur dans 1 minute...");
  try {
    // Nettoyage des anciens triggers
    const declencheurs = ScriptApp.getProjectTriggers();
    for (const declencheur of declencheurs) {
      if (declencheur.getHandlerFunction() === "traiterTacheArriereplan") {
        ScriptApp.deleteTrigger(declencheur);
      }
    }
    // Création du nouveau trigger
    ScriptApp.newTrigger("traiterTacheArriereplan")
      .timeBased()
      .after(60 * 1000)
      .create();
    console.log("Trigger programmé avec succès.");
  } catch (e) {
    console.error("Erreur fatale lors de la création du Trigger : " + e.message);
  }
}

/**
 * Fonction appelée par le trigger time-based.
 * Charge l'état sauvegardé, reprend la duplication,
 * et programme un nouveau trigger si nécessaire.
 */
function traiterTacheArriereplan() {
  console.log("Démarrage du script en arrière-plan via Trigger.");

  // Suppression immédiate du trigger courant
  try {
    const declencheurs = ScriptApp.getProjectTriggers();
    for (const declencheur of declencheurs) {
      if (declencheur.getHandlerFunction() === "traiterTacheArriereplan") {
        ScriptApp.deleteTrigger(declencheur);
      }
    }
  } catch (e) {
    console.warn("Impossible de supprimer le trigger actuel : " + e.message);
  }

  const tache = chargerEtat();
  if (!tache) {
    console.error("Aucune tâche trouvée dans PropertiesService. L'état a été perdu.");
    return;
  }

  try {
    const resultat = executerDuplication(tache, new Date().getTime(), true);

    if (resultat.termine) {
      console.log("Tâche terminée en arrière-plan. Nettoyage et envoi email.");
      supprimerEtat();
      sauvegarderHistorique(tache.nomSource, tache.nomDestination, tache.idDossierDest, resultat.statistiques);
      envoyerEmailFin_(tache.emailUtilisateur, tache.nomDestination, tache.idDossierDest, resultat.statistiques);
    } else {
      console.log("Tâche non terminée. Sauvegarde et relance du Trigger.");
      sauvegarderEtat(resultat.tache);
      programmerArriereplan();
    }
  } catch (e) {
    console.error("Crash inattendu dans traiterTacheArriereplan : " + e.message);
    envoyerEmailErreur_(tache, e);
  }
}

/**
 * Envoie un email de confirmation à la fin du traitement.
 *
 * @param {string} email Adresse email du destinataire
 * @param {string} nomDestination Nom du dossier de destination
 * @param {string} idDest ID du dossier de destination
 * @param {!Object} statistiques Statistiques de la duplication
 * @private
 */
function envoyerEmailFin_(email, nomDestination, idDest, statistiques) {
  if (!email) return;

  const url = `https://drive.google.com/drive/folders/${idDest}`;
  const objet = t('EMAIL_SUCCESS_SUBJECT', nomDestination);
  const nbFichiers = statistiques.fichiers || 0;
  const aErreurs = statistiques.erreurs && statistiques.erreurs.length > 0;

  // Fallback texte brut
  let corpsTexte = `${t('EMAIL_SUCCESS_HELLO')},\n\n`;
  corpsTexte += `${t('EMAIL_SUCCESS_BODY_1')}\n\n`;
  corpsTexte += `${t('EMAIL_SUCCESS_FOLDER', nomDestination)}\n`;
  corpsTexte += `${t('FOLDERS_CREATED')} : ${statistiques.dossiers}\n`;
  corpsTexte += `${t('FILES_COPIED')} : ${nbFichiers}\n\n`;
  if (aErreurs) {
    corpsTexte += t('EMAIL_SUCCESS_ERR', statistiques.erreurs.length) + "\n\n";
  }
  corpsTexte += t('EMAIL_SUCCESS_ACCESS', url);

  // Email HTML — Material Design
  const erreursHtml = aErreurs
    ? `<tr>
        <td style="padding:12px 16px;font-size:14px;color:#D93025;border-top:1px solid #E0E0E0;">
          ${t('EMAIL_SUCCESS_ERR_HTML', statistiques.erreurs.length)}
        </td>
       </tr>`
    : "";

  const corpsHtml = `
<!DOCTYPE html>
<html lang="fr">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background-color:#F5F5F5;font-family:'Google Sans','Roboto','Segoe UI',Arial,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F5F5F5;padding:24px 0;">
    <tr><td align="center">
      <table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background-color:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.12),0 1px 2px rgba(0,0,0,0.08);">

        <!-- Header -->
        <tr>
          <td style="background-color:#1A73E8;padding:32px 24px;text-align:center;">
            <div style="font-size:28px;margin-bottom:8px;">✅</div>
            <h1 style="margin:0;font-size:22px;font-weight:500;color:#FFFFFF;letter-spacing:0.15px;">
              ${t('DUPLICATION_FINISHED_TITLE')}
            </h1>
            <p style="margin:8px 0 0;font-size:14px;color:rgba(255,255,255,0.87);">
              ${t('EMAIL_SUCCESS_SUBTITLE')}
            </p>
          </td>
        </tr>

        <!-- Body -->
        <tr>
          <td style="padding:24px;">
            <p style="margin:0 0 20px;font-size:14px;line-height:1.6;color:#202124;">
              ${t('EMAIL_SUCCESS_HELLO')},
            </p>
            <p style="margin:0 0 24px;font-size:14px;line-height:1.6;color:#5F6368;">
              ${t('EMAIL_SUCCESS_BODY_2')}
            </p>

            <!-- Stats Card -->
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F8F9FA;border-radius:12px;overflow:hidden;margin-bottom:24px;">
              <tr>
                <td style="padding:16px;border-bottom:1px solid #E0E0E0;">
                  <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="font-size:12px;font-weight:500;color:#5F6368;text-transform:uppercase;letter-spacing:0.5px;">
                        ${t('FOLDER_CREATED')}
                      </td>
                    </tr>
                    <tr>
                      <td style="font-size:16px;font-weight:500;color:#202124;padding-top:4px;">
                        ${nomDestination}
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td style="padding:0;">
                  <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="50%" style="padding:16px;text-align:center;border-right:1px solid #E0E0E0;">
                        <div style="font-size:28px;font-weight:500;color:#1A73E8;">${statistiques.dossiers}</div>
                        <div style="font-size:12px;color:#5F6368;margin-top:4px;">${t('SUBFOLDERS')}</div>
                      </td>
                      <td width="50%" style="padding:16px;text-align:center;">
                        <div style="font-size:28px;font-weight:500;color:#1A73E8;">${nbFichiers}</div>
                        <div style="font-size:12px;color:#5F6368;margin-top:4px;">${t('FILES')}</div>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              ${erreursHtml}
            </table>

            <!-- CTA Button -->
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center" style="padding:8px 0 16px;">
                  <a href="${url}" target="_blank"
                     style="display:inline-block;padding:12px 32px;background-color:#1A73E8;color:#FFFFFF;
                            font-size:14px;font-weight:500;text-decoration:none;border-radius:24px;
                            letter-spacing:0.25px;box-shadow:0 1px 2px rgba(0,0,0,0.15);">
                    ${t('EMAIL_SUCCESS_OPEN_BTN')}
                  </a>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="padding:16px 24px;background-color:#F8F9FA;border-top:1px solid #E8EAED;text-align:center;">
            <p style="margin:0;font-size:12px;color:#9AA0A6;">
              ${VERSION} · ${AUTEUR}<br>
              ${t('EMAIL_FOOTER_DESC')}
            </p>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;

  try {
    MailApp.sendEmail({ to: email, subject: objet, body: corpsTexte, htmlBody: corpsHtml });
  } catch (e) {
    console.error("Erreur d'envoi email : " + e.message);
  }
}

/**
 * Envoie un email en cas de crash du processus d'arrière-plan.
 *
 * @param {!Object} tache État de la tâche au moment du crash
 * @param {!Error} erreur L'erreur survenue
 * @private
 */
function envoyerEmailErreur_(tache, erreur) {
  if (!tache || !tache.emailUtilisateur) return;

  const objet = t('EMAIL_FAILED_SUBJECT', tache.nomDestination);

  // Fallback texte brut
  const corpsTexte = `${t('EMAIL_SUCCESS_HELLO')},\n\n` +
    `${t('EMAIL_FAILED_BODY_1')}\n\n` +
    erreur.message + "\n\n" +
    `${t('EMAIL_FAILED_BODY_2')}`;

  // Email HTML — Material Design (thème erreur)
  const corpsHtml = `
<!DOCTYPE html>
<html lang="fr">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background-color:#F5F5F5;font-family:'Google Sans','Roboto','Segoe UI',Arial,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F5F5F5;padding:24px 0;">
    <tr><td align="center">
      <table role="presentation" width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;background-color:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.12),0 1px 2px rgba(0,0,0,0.08);">

        <!-- Header (erreur) -->
        <tr>
          <td style="background-color:#D93025;padding:32px 24px;text-align:center;">
            <div style="font-size:28px;margin-bottom:8px;">⚠️</div>
            <h1 style="margin:0;font-size:22px;font-weight:500;color:#FFFFFF;letter-spacing:0.15px;">
              ${t('EMAIL_FAILED_TITLE')}
            </h1>
            <p style="margin:8px 0 0;font-size:14px;color:rgba(255,255,255,0.87);">
              ${tache.nomDestination}
            </p>
          </td>
        </tr>

        <!-- Body -->
        <tr>
          <td style="padding:24px;">
            <p style="margin:0 0 20px;font-size:14px;line-height:1.6;color:#202124;">
              ${t('EMAIL_SUCCESS_HELLO')},
            </p>
            <p style="margin:0 0 24px;font-size:14px;line-height:1.6;color:#5F6368;">
              ${t('EMAIL_FAILED_BODY_1_HTML')}
            </p>

            <!-- Error Details -->
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#FEF7F6;border:1px solid #F5C6CB;border-radius:12px;overflow:hidden;margin-bottom:24px;">
              <tr>
                <td style="padding:16px;">
                  <div style="font-size:12px;font-weight:500;color:#D93025;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">
                    ${t('EMAIL_FAILED_DETAIL')}
                  </div>
                  <div style="font-size:14px;color:#202124;font-family:'Roboto Mono',monospace;word-break:break-word;">
                    ${erreur.message}
                  </div>
                </td>
              </tr>
            </table>

            <!-- Action -->
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F8F9FA;border-radius:12px;overflow:hidden;margin-bottom:24px;">
              <tr>
                <td style="padding:16px;">
                  <div style="font-size:12px;font-weight:500;color:#5F6368;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">
                    ${t('EMAIL_FAILED_ACTION_TITLE')}
                  </div>
                  <div style="font-size:14px;color:#202124;line-height:1.6;">
                    ${t('EMAIL_FAILED_ACTION_DESC')}
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Footer -->
        <tr>
          <td style="padding:16px 24px;background-color:#F8F9FA;border-top:1px solid #E8EAED;text-align:center;">
            <p style="margin:0;font-size:12px;color:#9AA0A6;">
              ${VERSION} · ${AUTEUR}<br>
              ${t('EMAIL_FOOTER_DESC')}
            </p>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>
</body>
</html>`;

  try {
    MailApp.sendEmail({ to: tache.emailUtilisateur, subject: objet, body: corpsTexte, htmlBody: corpsHtml });
  } catch (e) {
    console.error("Impossible d'envoyer l'email de crash : " + e.message);
  }
}

// ============================================================
// REPRISE MANUELLE
// ============================================================

/**
 * Reprend une duplication interrompue depuis le dernier état sauvegardé.
 *
 * @return {!Card} Carte de résultat, progression ou erreur
 */
function reprendreDuplication() {
  const tache = chargerEtat();
  if (!tache) return construireCarteErreur(t('NO_DUPLICATION_CORRUPTED'));

  const resultat = executerDuplication(tache, new Date().getTime());

  if (resultat.termine) {
    supprimerEtat();
    sauvegarderHistorique(tache.nomSource, tache.nomDestination, tache.idDossierDest, resultat.statistiques);
    return construireCarteResultat(tache.nomDestination, tache.idDossierDest, resultat.statistiques, tache.copierFichiers);
  } else {
    sauvegarderEtat(resultat.tache);
    if (tache.executerArriereplan) {
      programmerArriereplan();
      return construireCarteArriereplan(tache.idDossierDest);
    }
    return construireCarteProgression(resultat.tache);
  }
}

// ============================================================
// HISTORIQUE
// ============================================================

/**
 * Construit la section d'historique des duplications.
 * Affiche les MAX_HISTORIQUE dernières opérations avec liens vers les dossiers.
 *
 * @return {!CardSection} Section d'historique
 */
function construireSectionHistorique() {
  const section = CardService.newCardSection()
    .setHeader(t('HISTORY_HEADER'))
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0);

  const historique = chargerHistorique_();

  if (historique.length === 0) {
    return section.addWidget(CardService.newTextParagraph()
      .setText(
        `<font color="${COULEURS_MD.SUR_SURFACE_VARIANTE}">` +
        t('HISTORY_EMPTY') + "</font>"
      ));
  }

  historique.slice(0, MAX_HISTORIQUE).forEach(entree => {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel(entree.date)
      .setText(`${entree.nomSource} → ${entree.nomDestination}`)
      .setBottomLabel(`${entree.dossiers} ${t('FOLDERS_CREATED').toLowerCase()}, ${entree.fichiers} ${t('FILES_COPIED').toLowerCase()}`)
      .setWrapText(true)
      .setButton(CardService.newImageButton()
        .setIconUrl(ICONES_MD.OUVRIR_NOUVEAU)
        .setAltText(t('HISTORY_OPEN_ALT'))
        .setOpenLink(CardService.newOpenLink()
          .setUrl(`https://drive.google.com/drive/folders/${entree.idDest}`))));
  });

  section.addWidget(CardService.newDivider());
  section.addWidget(CardService.newTextButton()
    .setText(t('CLEAR_HISTORY_BTN'))
    .setOnClickAction(CardService.newAction().setFunctionName("effacerHistorique")));

  return section;
}

/**
 * Sauvegarde une entrée dans l'historique des duplications.
 *
 * @param {string} nomSource Nom du dossier source
 * @param {string} nomDestination Nom du dossier de destination
 * @param {string} idDest ID du dossier de destination
 * @param {!Object} statistiques Statistiques de la duplication
 */
function sauvegarderHistorique(nomSource, nomDestination, idDest, statistiques) {
  const historique = chargerHistorique_();

  historique.unshift({
    date: new Date().toLocaleString(Session.getActiveUserLocale() || "en-US"),
    nomSource,
    nomDestination,
    idDest,
    dossiers: statistiques.dossiers,
    fichiers: statistiques.fichiers || 0
  });

  PROPRIETES.setProperty("historique", JSON.stringify(historique.slice(0, MAX_HISTORIQUE)));
}

/**
 * Charge l'historique depuis PropertiesService avec gestion d'erreur.
 *
 * @return {!Array<!Object>} Tableau d'entrées d'historique
 * @private
 */
function chargerHistorique_() {
  const brut = PROPRIETES.getProperty("historique");
  if (!brut) return [];
  try {
    return JSON.parse(brut);
  } catch (e) {
    console.warn("Historique corrompu, réinitialisation.");
    PROPRIETES.deleteProperty("historique");
    return [];
  }
}

/**
 * Efface l'historique des duplications et affiche une notification.
 * Utilise setStateChanged(true) pour forcer le rafraîchissement
 * de la carte avec le contexte Drive correct.
 *
 * @return {!ActionResponse} Réponse avec notification et rafraîchissement
 */
function effacerHistorique() {
  PROPRIETES.deleteProperty("historique");
  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(t('HISTORY_CLEARED'))
    )
    .setStateChanged(true)
    .build();
}

// ============================================================
// CONSTRUCTEURS DE CARTES (UI)
// ============================================================

/**
 * Construit la section de pied de page avec version et auteur.
 *
 * @return {!CardSection} Section footer
 */
function construirePiedDePage() {
  return CardService.newCardSection()
    .setCollapsible(false)
    .addWidget(CardService.newTextParagraph()
      .setText(
        `<font color="${COULEURS_MD.SUR_SURFACE_VARIANTE}"><i>` +
        `${VERSION} · ${AUTEUR}` +
        "</i></font>"
      ));
}

/**
 * Construit la carte affichée pendant le traitement en arrière-plan.
 *
 * @param {string} idDossierDest ID du dossier en cours de construction
 * @return {!Card} Carte d'attente
 */
function construireCarteArriereplan(idDossierDest) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(t('BG_PROCESSING_TITLE'))
      .setSubtitle(t('BG_PROCESSING_SUB'))
      .setImageUrl(ICONES_MD.LOGO)
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(
          `<font color="${COULEURS_MD.PRIMAIRE}"><b>${t('BG_PROCESSING_MSG1')}</b></font>` +
          "<br><br>" + t('BG_PROCESSING_MSG2')
        ))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextButton()
        .setText(t('OPEN_FOLDER_BUILD'))
        .setOpenLink(CardService.newOpenLink()
          .setUrl(`https://drive.google.com/drive/folders/${idDossierDest}`))))
    .addSection(construirePiedDePage())
    .build();
}

/**
 * Construit la carte de progression pour une passe incomplète.
 *
 * @param {!Object} tache État de la tâche avec statistiques et file d'attente
 * @return {!Card} Carte de progression
 */
function construireCarteProgression(tache) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(t('PASS_INCOMPLETE_TITLE'))
      .setSubtitle(t('PASS_INCOMPLETE_SUB'))
      .setImageUrl(ICONES_MD.LOGO)
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(t('FOLDERS_CREATED')).setText(`${tache.statistiques.dossiers}`))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(t('FILES_COPIED')).setText(`${tache.statistiques.fichiers || 0}`))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(t('REMAINING_TO_PROCESS'))
        .setText(t('FOLDERS_IN_QUEUE', tache.fileAttente.length)))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextParagraph()
        .setText(
          `<font color="${COULEURS_MD.AVERTISSEMENT}">` +
          t('LARGE_TREE_CONTINUE') +
          "</font>"
        ))
      .addWidget(CardService.newTextButton()
        .setText(t('CONTINUE_BTN'))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName("reprendreDuplication"))))
    .addSection(construirePiedDePage())
    .build();
}

/**
 * Construit la carte de résultat finale (duplication terminée).
 *
 * @param {string} nomDestination Nom du dossier créé
 * @param {string} idDossierDest ID du dossier créé
 * @param {!Object} statistiques Statistiques finales { dossiers, fichiers, erreurs, ignores }
 * @param {boolean} copierFichiers true si les fichiers ont été copiés
 * @return {!Card} Carte de résultat
 */
function construireCarteResultat(nomDestination, idDossierDest, statistiques, copierFichiers) {
  const section = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel(t('FOLDER_CREATED')).setText(nomDestination).setWrapText(true))
    .addWidget(CardService.newDecoratedText()
      .setTopLabel(t('SUBFOLDERS_CREATED')).setText(`${statistiques.dossiers}`));

  if (copierFichiers) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel(t('FILES_COPIED')).setText(`${statistiques.fichiers || 0}`));
  }
  if (statistiques.ignores > 0) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel(t('FOLDERS_IGNORED')).setText(`${statistiques.ignores}`));
  }
  if (statistiques.erreurs && statistiques.erreurs.length > 0) {
    let texteErreur = `<font color="${COULEURS_MD.ERREUR}"><b>` +
      t('ERRORS_ENCOUNTERED', statistiques.erreurs.length) + `</b></font><br>`;
    statistiques.erreurs.slice(0, 5).forEach(err => {
      texteErreur += `• <i>${err}</i><br>`;
    });
    if (statistiques.erreurs.length > 5) {
      texteErreur += `<i>` + t('AND_OTHERS', statistiques.erreurs.length - 5) + `</i>`;
    }
    section.addWidget(CardService.newTextParagraph().setText(texteErreur));
  }

  section.addWidget(CardService.newDivider());
  section.addWidget(CardService.newTextButton()
    .setText(t('OPEN_CREATED_FOLDER'))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOpenLink(CardService.newOpenLink()
      .setUrl(`https://drive.google.com/drive/folders/${idDossierDest}`)));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(t('DUPLICATION_FINISHED_TITLE'))
      .setSubtitle(t('FOLDERS_CREATED_SUCCESS', statistiques.dossiers))
      .setImageUrl(ICONES_MD.LOGO)
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(section)
    .addSection(construirePiedDePage())
    .build();
}

/**
 * Construit une carte d'erreur avec message et bouton retour.
 *
 * @param {string} message Description de l'erreur
 * @return {!Card} Carte d'erreur
 */
function construireCarteErreur(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(t('ERROR_TITLE'))
      .setSubtitle(t('PROBLEM_OCCURRED'))
      .setImageUrl(ICONES_MD.LOGO)
      .setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`<font color="${COULEURS_MD.ERREUR}">${message}</font>`))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextButton()
        .setText(t('BACK_BTN'))
        .setOnClickAction(CardService.newAction()
          .setFunctionName("retourArriere"))))
    .addSection(construirePiedDePage())
    .build();
}

// ============================================================
// UTILITAIRES
// ============================================================

/**
 * Fonction utilitaire pour déclencher manuellement
 * le consentement OAuth lors du premier déploiement.
 */
function autoriserMaintenant() {
  const racine = DriveApp.getRootFolder();
  console.log("Autorisation complète OK — " + racine.getName());
}

/**
 * Synchronise les permissions d'un fichier ou dossier source vers la destination.
 * 
 * @param {string} idSource L'ID du fichier/dossier source.
 * @param {string} idDest L'ID du fichier/dossier de destination.
 * @private
 */
function synchroniserDroits_(idSource, idDest) {
  try {
    const permissionsListe = avecRetentative_(() => Drive.Permissions.list(idSource, {
      supportsAllDrives: true,
      fields: "permissions(emailAddress, role, type, deleted)"
    }));
    
    const permissions = permissionsListe.permissions;
    if (!permissions || permissions.length === 0) return;

    permissions.forEach(perm => {
      // On ignore les permissions supprimées ou sans email.
      if (perm.deleted || !perm.emailAddress) return;
      // On ignore le rôle "owner" car inhérent à la création
      if (perm.role === 'owner') return;

      try {
        avecRetentative_(() => Drive.Permissions.create({
          role: perm.role,
          type: perm.type,
          emailAddress: perm.emailAddress
        }, idDest, {
          supportsAllDrives: true,
          sendNotificationEmail: false
        }));
      } catch (err) {
        console.warn(`Impossible d'ajouter la permission pour ${perm.emailAddress} sur ${idDest} : ${err.message}`);
      }
    });
  } catch (e) {
    console.warn(`Impossible de lire les permissions du source ${idSource} : ${e.message}`);
  }
}
