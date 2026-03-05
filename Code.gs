// ============================================================
// DUPLIQUER UNE ARBORESCENCE - Google Drive Add-on
// ============================================================

const PROPS = PropertiesService.getUserProperties();
const MAX_EXECUTION_MS = 25 * 1000;
const VERSION = 'v2.1';
const AUTHOR = 'Fabrice FAUCHEUX';

// ============================================================
// POINT D'ENTRÉE
// ============================================================

function onDriveItemsSelected(e) {
  const items = e && e.drive && e.drive.selectedItems;
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Dupliquer une arborescence'));

  if (!items || items.length === 0) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('Sélectionnez un dossier dans Google Drive pour commencer.')));
    card.addSection(buildFooterSection());
    return card.build();
  }

  const item = items[0];
  if (item.mimeType !== 'application/vnd.google-apps.folder') {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('⚠️ Veuillez sélectionner un dossier (pas un fichier).')));
    card.addSection(buildFooterSection());
    return card.build();
  }

  card.addSection(CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel('Dossier source')
      .setText('📂 ' + item.title)
      .setWrapText(true)
      .setButton(CardService.newTextButton()
        .setText('Aperçu')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('previewFolder')
          .setParameters({ folderId: item.id, folderName: item.title })))));

  card.addSection(CardService.newCardSection()
    .setHeader('Options')
    .addWidget(CardService.newTextInput()
      .setFieldName('dest_name')
      .setTitle('Nom de la copie')
      .setValue('Copie de ' + item.title))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setTitle('Destination')
      .setFieldName('dest_location')
      .addItem('Même emplacement', 'same', true)
      .addItem('Racine de Mon Drive', 'root', false)
      .addItem('Autre dossier (coller son ID)', 'custom', false))
    .addWidget(CardService.newTextInput()
      .setFieldName('custom_dest_id')
      .setTitle('ID du dossier de destination')
      .setHint('Visible dans l\'URL Drive : .../folders/[ID]'))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName('copy_files')
      .addItem('Copier les fichiers en plus des dossiers', 'yes', false))
    .addWidget(CardService.newTextInput()
      .setFieldName('exclusions')
      .setTitle('Dossiers à ignorer')
      .setHint('Noms séparés par des virgules (optionnel)'))
    .addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName('use_regex')
      .addItem('Utiliser des expressions régulières pour les exclusions', 'yes', false)));

  const actionSection = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText('Lancer la duplication')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('startDuplication')
        .setParameters({ folderId: item.id, folderName: item.title })));

  if (loadJobState()) {
    actionSection.addWidget(CardService.newTextButton()
      .setText('▶️ Reprendre la duplication en cours')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('resumeDuplication')));
  }

  card.addSection(actionSection);
  card.addSection(buildHistorySection());
  card.addSection(buildFooterSection());

  return card.build();
}

function buildFooterSection() {
  return CardService.newCardSection()
    .setCollapsible(false)
    //.addWidget(CardService.newDivider())
    .addWidget(CardService.newTextParagraph()
      .setText('<font color="#888888"><i>' + VERSION + ' · ' + AUTHOR + '</i></font>'));
}

// ============================================================
// APERÇU LIMITÉ
// ============================================================

function previewFolder(e) {
  const folderId = e.parameters.folderId;
  const folderName = e.parameters.folderName;

  try {
    const folder = DriveApp.getFolderById(folderId);
    const stats = { folders: 0, files: 0, depth: 0, limitReached: false };
    countFolderContentsLimited(folder, stats, 0);

    const isLarge = stats.folders > 50 || stats.files > 200 || stats.limitReached;
    let limitWarning = stats.limitReached ? '\n\n⚠️ Limite d\'aperçu atteinte (>1000 éléments). L\'arborescence réelle est plus grande.' : '';

    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Aperçu : ' + folderName))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newDecoratedText()
          .setTopLabel('Sous-dossiers').setText(stats.folders + (stats.limitReached ? '+' : '')))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel('Fichiers').setText(stats.files + (stats.limitReached ? '+' : '')))
        .addWidget(CardService.newDecoratedText()
          .setTopLabel('Profondeur').setText(stats.depth + ' niveau(x)'))
        .addWidget(CardService.newTextParagraph()
          .setText((isLarge
            ? '⚠️ Arborescence volumineuse : la duplication se fera en plusieurs passes.'
            : '✅ Taille raisonnable, la duplication s\'effectuera en une seule passe.') + limitWarning))
        .addWidget(CardService.newTextButton()
          .setText('← Retour')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('onDriveItemsSelected'))))
      .addSection(buildFooterSection())
      .build();

  } catch (err) {
    return buildErrorCard('Impossible de lire le dossier : ' + err.message);
  }
}

function countFolderContentsLimited(folder, stats, depth) {
  if (stats.folders + stats.files > 1000) {
    stats.limitReached = true;
    return;
  }
  if (depth > stats.depth) stats.depth = depth;
  const subFolders = folder.getFolders();
  while (subFolders.hasNext() && !stats.limitReached) {
    stats.folders++;
    countFolderContentsLimited(subFolders.next(), stats, depth + 1);
  }
  const files = folder.getFiles();
  while (files.hasNext() && !stats.limitReached) { 
    files.next(); 
    stats.files++; 
    if (stats.folders + stats.files > 1000) stats.limitReached = true;
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
  PROPS.setProperty('job_chunks_count', chunksCount.toString());
  for(let i = 0; i < chunksCount; i++) {
    PROPS.setProperty('job_chunk_' + i, json.substring(i * chunkSize, (i + 1) * chunkSize));
  }
}

function loadJobState() {
  const countStr = PROPS.getProperty('job_chunks_count');
  if (!countStr) return null;
  const count = parseInt(countStr, 10);
  let json = '';
  for(let i = 0; i < count; i++) {
    json += PROPS.getProperty('job_chunk_' + i) || '';
  }
  try { return JSON.parse(json); } catch(e) { return null; }
}

function clearJobState() {
  const countStr = PROPS.getProperty('job_chunks_count');
  if (countStr) {
    const count = parseInt(countStr, 10);
    for(let i = 0; i < count; i++) PROPS.deleteProperty('job_chunk_' + i);
    PROPS.deleteProperty('job_chunks_count');
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

  const folderId = e.parameters.folderId;
  const folderName = e.parameters.folderName;
  const formInputs = e.formInputs || {};

  const destName = formInputs.dest_name ? formInputs.dest_name[0] : 'Copie de ' + folderName;
  const copyFiles = formInputs.copy_files && formInputs.copy_files[0] === 'yes';
  const destLocation = formInputs.dest_location ? formInputs.dest_location[0] : 'same';
  const customDestId = formInputs.custom_dest_id ? formInputs.custom_dest_id[0].trim() : '';
  const exclusionsRaw = formInputs.exclusions ? formInputs.exclusions[0] : '';
  const useRegex = formInputs.use_regex && formInputs.use_regex[0] === 'yes';
  
  const exclusions = exclusionsRaw.split(',').map(function(s) { return s.trim(); }).filter(Boolean);

  try {
    const sourceFolder = DriveApp.getFolderById(folderId);
    let parentFolderId;

    if (destLocation === 'custom' && customDestId) {
      try { DriveApp.getFolderById(customDestId); parentFolderId = customDestId; }
      catch(err) { return buildErrorCard('ID de destination invalide ou inaccessible.'); }
    } else if (destLocation === 'root') {
      parentFolderId = DriveApp.getRootFolder().getId();
    } else {
      const parents = sourceFolder.getParents();
      parentFolderId = parents.hasNext() ? parents.next().getId() : DriveApp.getRootFolder().getId();
    }

    const destFolder = withRetry(() => Drive.Files.create({
      name: destName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentFolderId]
    }, null, { supportsAllDrives: true }));

    const destFolderId = destFolder.id;

    const stats = { folders: 0, files: 0, errors: [], skipped: 0 };
    const job = {
      sourceFolderId: folderId,
      sourceName: folderName,
      destFolderId: destFolderId,
      destName: destName,
      copyFiles: copyFiles,
      exclusions: exclusions,
      useRegex: useRegex,
      stats: stats,
      queue: [[folderId, destFolderId]] 
    };

    const result = runDuplicationJob(job, new Date().getTime());

    if (result.completed) {
      clearJobState();
      saveToHistory(folderName, destName, destFolderId, result.stats);
      return buildResultCard(destName, destFolderId, result.stats, copyFiles);
    } else {
      saveJobState(result.job);
      return buildProgressCard(result.job);
    }

  } catch (err) {
    clearJobState();
    return buildErrorCard(err.message);
  }
}

function runDuplicationJob(job, startTime) {
  const queue = job.queue;

  function addError(msg) {
    if (job.stats.errors.length < 50) {
      job.stats.errors.push(msg.length > 100 ? msg.substring(0, 97) + '...' : msg);
    }
  }

  function isExcluded(name) {
    if (!job.exclusions || job.exclusions.length === 0) return false;
    for (let i = 0; i < job.exclusions.length; i++) {
      if (job.useRegex) {
        try {
          const regex = new RegExp(job.exclusions[i], 'i');
          if (regex.test(name)) return true;
        } catch(e) { /* ignore invalid regex */ }
      } else {
        if (name.toLowerCase().trim() === job.exclusions[i].toLowerCase()) return true;
      }
    }
    return false;
  }

  while (queue.length > 0) {
    if (new Date().getTime() - startTime > MAX_EXECUTION_MS) {
      job.queue = queue;
      return { completed: false, job: job };
    }

    const current = queue.pop();
    const currentSourceId = current[0];
    const currentDestId = current[1];
    let pageToken = null;

    do {
      try {
        const query = "'" + currentSourceId + "' in parents and trashed = false";
        const response = withRetry(() => Drive.Files.list({
          q: query,
          fields: "nextPageToken, files(id, name, mimeType)",
          pageSize: 1000,
          pageToken: pageToken,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true
        }));
        
        const children = response.files;

        if (children && children.length > 0) {
          for (let i = 0; i < children.length; i++) {
            const child = children[i];

            if (child.mimeType === 'application/vnd.google-apps.folder') {
              if (isExcluded(child.name)) {
                job.stats.skipped++;
                continue;
              }
              try {
                const newFolder = withRetry(() => Drive.Files.create({
                  name: child.name,
                  mimeType: 'application/vnd.google-apps.folder',
                  parents: [currentDestId]
                }, null, { supportsAllDrives: true }));

                job.stats.folders++;
                queue.push([child.id, newFolder.id]);
              } catch (errFolder) {
                addError('Création dossier "' + child.name + '" : ' + errFolder.message);
              }

            } else if (child.mimeType === 'application/vnd.google-apps.shortcut') {
              continue;
            } else {
              if (job.copyFiles) {
                try {
                  withRetry(() => Drive.Files.copy({
                    name: child.name,
                    parents: [currentDestId]
                  }, child.id, { supportsAllDrives: true }));
                  job.stats.files++;
                } catch (errFile) {
                  addError('Copie fichier "' + child.name + '" : ' + errFile.message);
                }
              }
            }
          }
        }
        pageToken = response.nextPageToken;
      } catch (e) {
        addError('Lecture du dossier source échouée : ' + e.message);
        break;
      }
    } while (pageToken);
  }

  return { completed: true, stats: job.stats };
}

// ============================================================
// REPRISE MANUELLE
// ============================================================

function resumeDuplication(e) {
  const job = loadJobState();
  if (!job) return buildErrorCard('Aucune duplication en cours ou données corrompues.');

  const result = runDuplicationJob(job, new Date().getTime());

  if (result.completed) {
    clearJobState();
    saveToHistory(job.sourceName, job.destName, job.destFolderId, result.stats);
    return buildResultCard(job.destName, job.destFolderId, result.stats, job.copyFiles);
  } else {
    saveJobState(result.job);
    return buildProgressCard(result.job);
  }
}

// ============================================================
// HISTORIQUE ET UTILITAIRES DE CARTES (INCHANGÉS OU PRESQUE)
// ============================================================

function buildHistorySection() {
  const section = CardService.newCardSection()
    .setHeader('Historique')
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0);
  const historyRaw = PROPS.getProperty('history');
  const history = historyRaw ? JSON.parse(historyRaw) : [];

  if (history.length === 0) {
    return section.addWidget(CardService.newTextParagraph()
      .setText('Aucune duplication effectuée pour l\'instant.'));
  }

  history.slice(0, 10).forEach(function(entry) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel(entry.date)
      .setText(entry.sourceName + ' → ' + entry.destName)
      .setBottomLabel(entry.folders + ' dossiers, ' + entry.files + ' fichiers')
      .setWrapText(true)
      .setButton(CardService.newImageButton()
        .setIconUrl('https://www.gstatic.com/images/icons/material/system/1x/open_in_new_googblue_18dp.png')
        .setAltText('Ouvrir')
        .setOpenLink(CardService.newOpenLink()
          .setUrl('https://drive.google.com/drive/folders/' + entry.destId))));
  });

  section.addWidget(CardService.newTextButton()
    .setText('Effacer l\'historique')
    .setOnClickAction(CardService.newAction().setFunctionName('clearHistory')));

  return section;
}

function saveToHistory(sourceName, destName, destId, stats) {
  const historyRaw = PROPS.getProperty('history');
  const history = historyRaw ? JSON.parse(historyRaw) : [];
  history.unshift({
    date: new Date().toLocaleString('fr-FR'),
    sourceName: sourceName,
    destName: destName,
    destId: destId,
    folders: stats.folders,
    files: stats.files || 0
  });
  PROPS.setProperty('history', JSON.stringify(history.slice(0, 10)));
}

function clearHistory(e) {
  PROPS.deleteProperty('history');
  const updatedCard = onDriveItemsSelected(e);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Historique effacé.'))
    .setNavigation(CardService.newNavigation().updateCard(updatedCard))
    .build();
}

function buildProgressCard(job) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('⏳ Passe incomplète'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel('Dossiers créés').setText('' + job.stats.folders))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel('Fichiers copiés').setText('' + (job.stats.files || 0)))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel('Restant à traiter').setText(job.queue.length + ' dossier(s) en file d\'attente'))
      .addWidget(CardService.newTextParagraph()
        .setText('L\'arborescence est volumineuse. Cliquez sur "Continuer" pour traiter la passe suivante.'))
      .addWidget(CardService.newTextButton()
        .setText('▶️ Continuer la duplication')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('resumeDuplication'))))
    .addSection(buildFooterSection())
    .build();
}

function buildResultCard(destName, destFolderId, stats, copyFiles) {
  const section = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel('Dossier créé').setText(destName).setWrapText(true))
    .addWidget(CardService.newDecoratedText()
      .setTopLabel('Sous-dossiers créés').setText('' + stats.folders));

  if (copyFiles) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel('Fichiers copiés').setText('' + (stats.files || 0)));
  }
  if (stats.skipped > 0) {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel('Dossiers ignorés').setText('' + stats.skipped));
  }
  if (stats.errors && stats.errors.length > 0) {
    let errorText = '<font color="#BA0000"><b>⚠️ ' + stats.errors.length + ' erreur(s) rencontrée(s) :</b></font><br>';
    stats.errors.slice(0, 5).forEach(function(err) {
      errorText += '• <i>' + err + '</i><br>';
    });
    if (stats.errors.length > 5) {
      errorText += '<i>...et ' + (stats.errors.length - 5) + ' autre(s).</i>';
    }
    section.addWidget(CardService.newTextParagraph().setText(errorText));
  }

  section.addWidget(CardService.newDivider());
  section.addWidget(CardService.newTextButton()
    .setText('Ouvrir le dossier créé')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOpenLink(CardService.newOpenLink()
      .setUrl('https://drive.google.com/drive/folders/' + destFolderId)));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('✅ Duplication terminée !'))
    .addSection(section)
    .addSection(buildFooterSection())
    .build();
}

function buildErrorCard(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('❌ Erreur'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(message))
      .addWidget(CardService.newTextButton()
        .setText('← Retour')
        .setOnClickAction(CardService.newAction().setFunctionName('onDriveItemsSelected'))))
    .addSection(buildFooterSection())
    .build();
}

function authorizeNow() {
  const root = DriveApp.getRootFolder();
  const testFolder = root.createFolder('__test_autorisation__');
  testFolder.setTrashed(true);
  console.log('Autorisation complète OK');
}
