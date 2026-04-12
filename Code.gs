// ============================================================
// DUPLIQUER UNE ARBORESCENCE — Google Drive Add-on
// ============================================================
// Auteur  : Fabrice FAUCHEUX
// Version : v3.1
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
//    d'attente (BFS) pour éviter tout risque de récursion.
//
// GESTION DU TIMEOUT
// ------------------
// Google Apps Script impose une limite d'exécution de 6 min.
// Ce script utilise deux seuils de sécurité :
//   - 25 s  en mode interactif (l'utilisateur attend)
//   - 4 min en mode arrière-plan (trigger automatique)
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
// processBackgroundJob() toutes les ~1 minute jusqu'à
// complétion. Un email de confirmation est envoyé à
// l'utilisateur à la fin du traitement.
//
// STRUCTURE DU CODE
// -----------------
//   onDriveItemsSelected()     Point d'entrée principal (UI)
//   previewFolder()            Aperçu rapide de l'arborescence
//   startDuplication()         Initialisation et lancement
//   runDuplicationJob()        Moteur de duplication (BFS)
//   scheduleBackgroundJob()    Création du trigger de reprise
//   processBackgroundJob()     Exécution en arrière-plan
//   resumeDuplication()        Reprise manuelle depuis l'UI
//   saveJobState/loadJobState  Persistance par chunks
//   buildXxxCard()             Constructeurs de cartes UI
//   saveToHistory()            Historique des duplications
//
// PERMISSIONS REQUISES
// --------------------
//   drive                          Lecture/écriture Drive
//   drive.addons.metadata.readonly Métadonnées Add-on
//   script.scriptapp               Gestion des triggers
//   script.send_mail               Envoi d'email de fin
//   userinfo.email                 Récupération de l'email
//
// ============================================================

const PROPS = PropertiesService.getUserProperties();
const MAX_EXECUTION_MS_UI = 25 * 1000;  // 25 secondes quand l'utilisateur attend
const MAX_EXECUTION_MS_BG = 250 * 1000; // 4 minutes quand ça tourne en arrière-plan
const VERSION = "v3.1";
const AUTHOR = "Fabrice FAUCHEUX";

// ============================================================
// POINT D'ENTRÉE
// ============================================================

function onDriveItemsSelected(e) {
  const items = e && e.drive && e.drive.selectedItems;
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle("Dupliquer une arborescence"));

  if (!items || items.length === 0) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText("Sélectionnez un dossier dans Google Drive pour commencer.")));
    card.addSection(buildFooterSection());
    return card.build();
  }

  const item = items[0];
  if (item.mimeType !== "application/vnd.google-apps.folder") {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText("⚠️ Veuillez sélectionner un dossier (pas un fichier).")));
    card.addSection(buildFooterSection());
    return card.build();
  }

  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel("Dossier source")
      .setText(`📂 ${item.title}`)
      .setWrapText(true)
      .setButton(CardService.newTextButton()
        .setText("Aperçu")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("previewFolder")
          .setParameters({ folderId: item.id, folderName: item.title })))));

  card.addSection(CardService.newCardSection()
    .setHeader("Options")
    .addWidget(CardService.newTextInput()
      .setFieldName("dest_name")
      .setTitle("Nom de la copie")
      .setValue(`Copie de ${item.title}`))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setTitle("Destination")
      .setFieldName("dest_location")
      .addItem("Même emplacement", "same", true)
      .addItem("Racine de Mon Drive", "root", false)
      .addItem("Autre dossier (coller son ID)", "custom", false))
    .addWidget(CardService.newTextInput()
      .setFieldName("custom_dest_id")
      .setTitle("ID du dossier de destination")
      .setHint("Visible dans l'URL Drive : .../folders/[ID]"))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("copy_files")
      .addItem("Copier les fichiers en plus des dossiers", "yes", false))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("run_background")
      .addItem("Automatiser en arrière-plan (recommandé)", "yes", true))
    .addWidget(CardService.newTextInput()
      .setFieldName("exclusions")
      .setTitle("Dossiers à ignorer")
      .setHint("Noms séparés par des virgules (optionnel)"))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("use_regex")
      .addItem("Utiliser des expressions régulières pour les exclusions", "yes", false)));

  const actionSection = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText("Lancer la duplication")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("startDuplication")
        .setParameters({ folderId: item.id, folderName: item.title })));

  if (loadJobState()) {
    actionSection.addWidget(CardService.newTextButton()
      .setText("▶️ Reprendre la duplication en cours")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("resumeDuplication")));
  }

  card.addSection(actionSection);
  card.addSection(buildHistorySection());
  card.addSection(buildFooterSection());

  return card.build();
}

function buildFooterSection() {
  return CardService.newCardSection()
    .setCollapsible(false)
    .addWidget(CardService.newTextParagraph()
      .setText(`<font color="#888888"><i>${VERSION} · ${AUTHOR}</i></font>`));
}

// ============================================================
// APERÇU LIMITÉ
// ============================================================

function previewFolder(e) {
  const { folderId, folderName } = e.parameters;

  try {
    const stats = { folders: 0, files: 0, depth: 0, limitReached: false };
    
    countFolderContentsLimitedAPI(folderId, stats, 0);

    const isLarge = stats.folders > 50 || stats.files > 200 || stats.limitReached;
    
    let limitWarning = "";
    if (stats.limitReached) {
      limitWarning = "\n\n⚠️ Limite d'aperçu atteinte (>1000 éléments). L'arborescence réelle est plus grande.";
    }

    let messageTaille = "✅ Taille raisonnable, la duplication s'effectuera en une seule passe.";
    if (isLarge) {
      messageTaille = "⚠️ Arborescence volumineuse : la duplication se fera en plusieurs passes.";
    }

    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle(`Aperçu : ${folderName}`))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newDecoratedText()
          .setTopLabel("Sous-dossiers").setText(`${stats.folders}${stats.limitReached ? "+" : ""}`))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel("Fichiers").setText(`${stats.files}${stats.limitReached ? "+" : ""}`))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel("Profondeur").setText(`${stats.depth} niveau(x)`))
        .addWidget(CardService.newTextParagraph()
          .setText(messageTaille + limitWarning))
        .addWidget(CardService.newTextButton()
          .setText("← Retour")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("onDriveItemsSelected"))))
      .addSection(buildFooterSection())
      .build();

  } catch (err) {
    return buildErrorCard(`Impossible de lire le dossier : ${err.message}`);
  }
}

function countFolderContentsLimitedAPI(folderId, stats, depth) {
  if (stats.folders + stats.files > 1000 || stats.limitReached) {
    stats.limitReached = true; 
    return;
  }
  if (depth > stats.depth) {
    stats.depth = depth;
  }

  let pageToken = null;
  const subFolders = [];

  do {
    const response = withRetry(() => Drive.Files.list({
      q: `'${folderId}' in parents and trashed = false`,
      fields: "nextPageToken, files(id, mimeType)",
      pageSize: 1000,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    }));
    
    const items = response.files || [];
    for (const item of items) {
      if (item.mimeType === "application/vnd.google-apps.folder") {
        stats.folders++;
        subFolders.push(item.id);
      } else if (item.mimeType !== "application/vnd.google-apps.shortcut") {
        stats.files++;
      }
    }
    if (stats.folders + stats.files > 1000) {
      stats.limitReached = true;
    }
    pageToken = response.nextPageToken;
  } while (pageToken && !stats.limitReached);

  for (const subFolderId of subFolders) {
    if (stats.limitReached) break;
    countFolderContentsLimitedAPI(subFolderId, stats, depth + 1);
  }
}

// ============================================================
// GESTION DU STOCKAGE (CHUNKING) & RETRY
// ============================================================

function saveJobState(job) {
  clearJobState();
  const json = JSON.stringify(job);
  const chunkSize = 8000;
  const chunksCount = Math.ceil(json.length / chunkSize);
  PROPS.setProperty("job_chunks_count", chunksCount.toString());
  
  for (let i = 0; i < chunksCount; i++) {
    PROPS.setProperty(`job_chunk_${i}`, json.substring(i * chunkSize, (i + 1) * chunkSize));
  }
}

function loadJobState() {
  const countStr = PROPS.getProperty("job_chunks_count");
  if (!countStr) return null;
  const count = parseInt(countStr, 10);
  
  let json = "";
  for (let i = 0; i < count; i++) {
    json += PROPS.getProperty(`job_chunk_${i}`) || "";
  }
  try { return JSON.parse(json); } catch(e) { return null; }
}

function clearJobState() {
  const countStr = PROPS.getProperty("job_chunks_count");
  if (countStr) {
    const count = parseInt(countStr, 10);
    for (let i = 0; i < count; i++) {
      PROPS.deleteProperty(`job_chunk_${i}`);
    }
    PROPS.deleteProperty("job_chunks_count");
  }
}

function withRetry(action, maxRetries = 3) {
  let attempt = 0;
  while (attempt < maxRetries) {
    try {
      return action();
    } catch (e) {
      attempt++;
      if (attempt >= maxRetries) throw e;
      Utilities.sleep(Math.pow(2, attempt) * 1000 + Math.round(Math.random() * 500));
    }
  }
}

// ============================================================
// LANCEMENT ET DUPLICATION
// ============================================================

function startDuplication(e) {
  clearJobState();

  const { folderId, folderName } = e.parameters;
  const formInputs = e.formInputs || {};

  // Remplacement des ?. et ?? par des vérifications ES6 classiques
  const destName = (formInputs.dest_name && formInputs.dest_name[0]) ? formInputs.dest_name[0] : `Copie de ${folderName}`;
  const copyFiles = formInputs.copy_files && formInputs.copy_files[0] === "yes";
  const runBackground = formInputs.run_background && formInputs.run_background[0] === "yes";
  const destLocation = (formInputs.dest_location && formInputs.dest_location[0]) ? formInputs.dest_location[0] : "same";
  const customDestId = (formInputs.custom_dest_id && formInputs.custom_dest_id[0]) ? formInputs.custom_dest_id[0].trim() : "";
  const exclusionsRaw = (formInputs.exclusions && formInputs.exclusions[0]) ? formInputs.exclusions[0] : "";
  const useRegex = formInputs.use_regex && formInputs.use_regex[0] === "yes";
  
  const exclusions = exclusionsRaw.split(",").map(s => s.trim()).filter(Boolean);

  try {
    const sourceFolder = DriveApp.getFolderById(folderId);
    let parentFolderId;

    if (destLocation === "custom" && customDestId) {
      try { 
        DriveApp.getFolderById(customDestId); 
        parentFolderId = customDestId; 
      } catch (err) { 
        return buildErrorCard("ID de destination invalide ou inaccessible."); 
      }
    } else if (destLocation === "root") {
      parentFolderId = DriveApp.getRootFolder().getId();
    } else {
      const parents = sourceFolder.getParents();
      parentFolderId = parents.hasNext() ? parents.next().getId() : DriveApp.getRootFolder().getId();
    }

    const destFolder = withRetry(() => Drive.Files.create({
      name: destName,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentFolderId]
    }, null, { supportsAllDrives: true }));

    const stats = { folders: 0, files: 0, errors: [], skipped: 0 };
    const job = {
      sourceFolderId: folderId,
      sourceName: folderName,
      destFolderId: destFolder.id,
      destName,
      copyFiles,
      runBackground,
      userEmail: Session.getEffectiveUser().getEmail(),
      exclusions,
      useRegex,
      stats,
      queue: [[folderId, destFolder.id]] 
    };

    const result = runDuplicationJob(job, new Date().getTime());

    if (result.completed) {
      clearJobState();
      saveToHistory(folderName, destName, destFolder.id, result.stats);
      return buildResultCard(destName, destFolder.id, result.stats, copyFiles);
    } else {
      saveJobState(result.job);
      if (runBackground) {
        scheduleBackgroundJob();
        return buildBackgroundCard(destFolder.id);
      } else {
        return buildProgressCard(result.job);
      }
    }

  } catch (err) {
    clearJobState();
    return buildErrorCard(err.message);
  }
}

function runDuplicationJob(job, startTime, isBackground = false) {
  const queue = job.queue;
  const compiledExclusions = [];
  const maxExecutionTime = isBackground ? MAX_EXECUTION_MS_BG : MAX_EXECUTION_MS_UI;
  
  console.log(`▶️ Début passe duplication. File d'attente : ${queue.length} dossiers. Mode BG : ${isBackground}`);

  if (job.exclusions && job.exclusions.length > 0) {
    for (const exclusion of job.exclusions) {
      if (job.useRegex) {
        try { compiledExclusions.push(new RegExp(exclusion, "i")); } catch (e) { }
      } else {
        compiledExclusions.push(exclusion.toLowerCase());
      }
    }
  }

  const addError = (msg) => {
    console.error(`❌ Erreur: ${msg}`);
    if (job.stats.errors.length < 50) {
      job.stats.errors.push(msg.length > 100 ? `${msg.substring(0, 97)}...` : msg);
    }
  };

  const isExcluded = (name) => {
    if (compiledExclusions.length === 0) return false;
    for (const compiled of compiledExclusions) {
      if (job.useRegex && compiled instanceof RegExp) {
        if (compiled.test(name)) return true;
      } else {
        if (name.toLowerCase().trim() === compiled) return true;
      }
    }
    return false;
  };

  while (queue.length > 0) {
    if (new Date().getTime() - startTime > maxExecutionTime) {
      job.queue = queue;
      console.log(`⏱️ Temps limite atteint (${maxExecutionTime}ms). Arrêt temporaire. Reste : ${queue.length} dossiers.`);
      return { completed: false, job };
    }

    const [currentSourceId, currentDestId] = queue.pop();
    let pageToken = null;

    do {
      try {
        const query = `'${currentSourceId}' in parents and trashed = false`;
        const response = withRetry(() => Drive.Files.list({
          q: query,
          fields: "nextPageToken, files(id, name, mimeType, description, folderColorRgb)",
          pageSize: 1000,
          pageToken: pageToken,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true
        }));
        
        const children = response.files || [];

        for (const child of children) {
          if (child.mimeType === "application/vnd.google-apps.folder") {
            if (isExcluded(child.name)) {
              job.stats.skipped++;
              continue;
            }
            try {
              const folderMetadata = {
                name: child.name,
                mimeType: "application/vnd.google-apps.folder",
                parents: [currentDestId]
              };
              if (child.description) folderMetadata.description = child.description;
              if (child.folderColorRgb) folderMetadata.folderColorRgb = child.folderColorRgb;

              const newFolder = withRetry(() => Drive.Files.create(folderMetadata, null, { supportsAllDrives: true }));

              job.stats.folders++;
              queue.push([child.id, newFolder.id]);
            } catch (errFolder) {
              addError(`Création dossier "${child.name}" : ${errFolder.message}`);
            }

          } else if (child.mimeType === "application/vnd.google-apps.shortcut") {
            continue;
          } else {
            if (job.copyFiles) {
              try {
                const fileMetadata = { parents: [currentDestId], name: child.name };
                if (child.description) fileMetadata.description = child.description;
                
                withRetry(() => Drive.Files.copy(fileMetadata, child.id, { supportsAllDrives: true }));
                
                job.stats.files++;
              } catch (errFile) {
                addError(`Copie fichier "${child.name}" : ${errFile.message}`);
              }
            }
          }
        }
        pageToken = response.nextPageToken;
      } catch (e) {
        addError(`Lecture du dossier source échouée : ${e.message}`);
        break;
      }
    } while (pageToken);
  }

  console.log(`✅ File d'attente vide. Duplication terminée !`);
  return { completed: true, stats: job.stats };
}

// ============================================================
// SYSTÈME D'ARRIÈRE-PLAN (TRIGGERS)
// ============================================================

// ============================================================
// SYSTÈME D'ARRIÈRE-PLAN (TRIGGERS) ET LOGS SÉCURISÉS
// ============================================================

function scheduleBackgroundJob() {
  console.log("🕒 Programmation d'un déclencheur (Trigger) dans 1 minute...");
  try {
    // Nettoyage agressif de tous les anciens triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "processBackgroundJob") {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    // Création du nouveau
    ScriptApp.newTrigger("processBackgroundJob").timeBased().after(60 * 1000).create();
    console.log("✅ Trigger programmé avec succès.");
  } catch (e) {
    console.error(`💥 Erreur fatale lors de la création du Trigger : ${e.message}`);
  }
}

function processBackgroundJob() {
  console.log("🚀 Démarrage du script en arrière-plan via Trigger.");
  
  // Suppression immédiate du trigger qui vient de se lancer
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "processBackgroundJob") {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  } catch (e) {
    console.warn("Impossible de supprimer le trigger actuel, on continue quand même.");
  }

  const job = loadJobState();
  if (!job) {
    console.error("❌ Aucun Job trouvé dans PropertiesService. L'état a été perdu.");
    return;
  }

  try {
    // On précise "true" pour dire qu'on est en arrière-plan (4 minutes de temps max)
    const result = runDuplicationJob(job, new Date().getTime(), true);

    if (result.completed) {
      console.log("🎉 Job terminé en arrière-plan ! Nettoyage et envoi email.");
      clearJobState();
      saveToHistory(job.sourceName, job.destName, job.destFolderId, result.stats);
      sendCompletionEmail(job.userEmail, job.destName, job.destFolderId, result.stats);
    } else {
      console.log("⏸️ Job non terminé, sauvegarde de l'état et relance du Trigger.");
      saveJobState(result.job);
      scheduleBackgroundJob();
    }
  } catch (e) {
    console.error(`💥 Crash inattendu dans processBackgroundJob : ${e.message}`);
    // Envoi d'un email de crash si possible
    if (job && job.userEmail) {
      MailApp.sendEmail({
        to: job.userEmail,
        subject: `❌ Échec de la duplication de ${job.destName}`,
        body: `Bonjour,\n\nLa duplication en arrière-plan s'est arrêtée suite à une erreur inattendue :\n\n${e.message}\n\nVous pouvez reprendre manuellement via l'interface du Add-on.`
      });
    }
  }
}

function sendCompletionEmail(email, destName, destId, stats) {
  if (!email) return; 
  
  const url = `https://drive.google.com/drive/folders/${destId}`;
  const subject = `✅ Duplication terminée : ${destName}`;
  
  let body = `Bonjour,\n\nLa duplication de votre arborescence est terminée avec succès.\n\n`;
  body += `Dossier : ${destName}\n`;
  body += `Dossiers créés : ${stats.folders}\n`;
  body += `Fichiers copiés : ${stats.files || 0}\n\n`;
  
  if (stats.errors && stats.errors.length > 0) {
    body += `⚠️ ${stats.errors.length} erreur(s) rencontrée(s) (voir le script pour les détails).\n\n`;
  }
  
  body += `Vous pouvez y accéder ici : ${url}`;

  try {
    MailApp.sendEmail({ to: email, subject, body });
  } catch(e) { 
    console.error("Erreur email:", e); 
  }
}

function buildBackgroundCard(destFolderId) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("🚀 En cours de traitement..."))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText("L'arborescence est volumineuse. **Le script a pris le relais en arrière-plan.**\n\nVous pouvez fermer cet outil. Vous recevrez un email dès que la copie sera 100% terminée."))
      .addWidget(CardService.newTextButton()
        .setText("Ouvrir le dossier (en construction)")
        .setOpenLink(CardService.newOpenLink()
          .setUrl(`https://drive.google.com/drive/folders/${destFolderId}`))))
    .addSection(buildFooterSection())
    .build();
}

// ============================================================
// REPRISE MANUELLE
// ============================================================

function resumeDuplication() {
  const job = loadJobState();
  if (!job) return buildErrorCard("Aucune duplication en cours ou données corrompues.");

  const result = runDuplicationJob(job, new Date().getTime());

  if (result.completed) {
    clearJobState();
    saveToHistory(job.sourceName, job.destName, job.destFolderId, result.stats);
    return buildResultCard(job.destName, job.destFolderId, result.stats, job.copyFiles);
  } else {
    saveJobState(result.job);
    if (job.runBackground) {
      scheduleBackgroundJob();
      return buildBackgroundCard(job.destFolderId);
    }
    return buildProgressCard(result.job);
  }
}

// ============================================================
// HISTORIQUE ET UTILITAIRES DE CARTES 
// ============================================================

function buildHistorySection() {
  const section = CardService.newCardSection()
    .setHeader("Historique")
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0);
  
  const historyRaw = PROPS.getProperty("history");
  const history = historyRaw ? JSON.parse(historyRaw) : [];

  if (history.length === 0) {
    return section.addWidget(CardService.newTextParagraph()
      .setText("Aucune duplication effectuée pour l'instant."));
  }

  history.slice(0, 10).forEach(entry => {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel(entry.date)
      .setText(`${entry.sourceName} → ${entry.destName}`)
      .setBottomLabel(`${entry.folders} dossiers, ${entry.files} fichiers`)
      .setWrapText(true)
      .setButton(CardService.newImageButton()
        .setIconUrl("https://www.gstatic.com/images/icons/material/system/1x/open_in_new_googblue_18dp.png")
        .setAltText("Ouvrir")
        .setOpenLink(CardService.newOpenLink()
          .setUrl(`https://drive.google.com/drive/folders/${entry.destId}`))));
  });

  section.addWidget(CardService.newTextButton()
    .setText("Effacer l'historique")
    .setOnClickAction(CardService.newAction().setFunctionName("clearHistory")));

  return section;
}

function saveToHistory(sourceName, destName, destId, stats) {
  const historyRaw = PROPS.getProperty("history");
  const history = historyRaw ? JSON.parse(historyRaw) : [];
  
  history.unshift({
    date: new Date().toLocaleString("fr-FR"),
    sourceName,
    destName,
    destId,
    folders: stats.folders,
    files: stats.files || 0
  });
  
  PROPS.setProperty("history", JSON.stringify(history.slice(0, 10)));
}

function clearHistory(e) {
  PROPS.deleteProperty("history");
  const updatedCard = onDriveItemsSelected(e);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText("Historique effacé."))
    .setNavigation(CardService.newNavigation().updateCard(updatedCard))
    .build();
}

function buildProgressCard(job) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("⏳ Passe incomplète"))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel("Dossiers créés").setText(`${job.stats.folders}`))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel("Fichiers copiés").setText(`${job.stats.files || 0}`))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel("Restant à traiter").setText(`${job.queue.length} dossier(s) en file d'attente`))
      .addWidget(CardService.newTextParagraph()
        .setText("L'arborescence est volumineuse. Cliquez sur Continuer pour traiter la passe suivante."))
      .addWidget(CardService.newTextButton()
        .setText("▶️ Continuer la duplication")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName("resumeDuplication"))))
    .addSection(buildFooterSection())
    .build();
}

function buildResultCard(destName, destFolderId, stats, copyFiles) {
  const section = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel("Dossier créé").setText(destName).setWrapText(true))
    .addWidget(CardService.newDecoratedText()
      .setTopLabel("Sous-dossiers créés").setText(`${stats.folders}`));

  if (copyFiles) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel("Fichiers copiés").setText(`${stats.files || 0}`));
  }
  if (stats.skipped > 0) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel("Dossiers ignorés").setText(`${stats.skipped}`));
  }
  if (stats.errors && stats.errors.length > 0) {
    let errorText = `<font color="#BA0000"><b>⚠️ ${stats.errors.length} erreur(s) rencontrée(s) :</b></font><br>`;
    stats.errors.slice(0, 5).forEach(err => {
      errorText += `• <i>${err}</i><br>`;
    });
    if (stats.errors.length > 5) {
      errorText += `<i>...et ${stats.errors.length - 5} autre(s).</i>`;
    }
    section.addWidget(CardService.newTextParagraph().setText(errorText));
  }

  section.addWidget(CardService.newDivider());
  section.addWidget(CardService.newTextButton()
    .setText("Ouvrir le dossier créé")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOpenLink(CardService.newOpenLink()
      .setUrl(`https://drive.google.com/drive/folders/${destFolderId}`)));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("✅ Duplication terminée !"))
    .addSection(section)
    .addSection(buildFooterSection())
    .build();
}

function buildErrorCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("❌ Erreur"))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(message))
      .addWidget(CardService.newTextButton()
        .setText("← Retour")
        .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected"))))
    .addSection(buildFooterSection())
    .build();
}

function authorizeNow() {
  const root = DriveApp.getRootFolder();
  console.log("Autorisation complète OK");
}
