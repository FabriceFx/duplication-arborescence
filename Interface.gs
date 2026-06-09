/**
 * ============================================================================
 * MODULE 4 : INTERFACE GRAPHIQUE (CARD SERVICE) ET COMMUNICATIONS
 * ============================================================================
 */

function onDriveItemsSelected(e) {
  const elements = e?.drive?.selectedItems;
  const carte = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(t('TITLE'))
      .setSubtitle(t('SUBTITLE'))
      .setImageUrl(CONFIG.ICONES.LOGO)
      .setImageStyle(CardService.ImageStyle.CIRCLE));

  if (!elements || elements.length === 0) {
    carte.addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph()
      .setText(`<font color="${CONFIG.COULEURS.SUR_SURFACE_VARIANTE}">${t('SELECT_FOLDER_PROMPT')}</font>`)));
    carte.addSection(construirePiedDePage());
    return carte.build();
  }

  const item = elements[0];
  if (item.mimeType !== CONFIG.MIME_DOSSIER) {
    carte.addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph()
      .setText(`<font color="${CONFIG.COULEURS.AVERTISSEMENT}"><b>${t('WARNING')}</b></font><br>${t('NOT_A_FOLDER')}`)));
    carte.addSection(construirePiedDePage());
    return carte.build();
  }

  // Source Section
  carte.addSection(CardService.newCardSection().addWidget(CardService.newDecoratedText()
    .setTopLabel(t('SOURCE_FOLDER')).setText(item.title).setWrapText(true)
    .setButton(CardService.newTextButton().setText(t('PREVIEW_BTN'))
      .setOnClickAction(CardService.newAction().setFunctionName("apercuDossier")
        .setParameters({ idDossier: item.id, nomDossier: item.title })))));

  // Setup Section
  carte.addSection(CardService.newCardSection()
    .setHeader(t('CONFIG_HEADER'))
    .addWidget(CardService.newTextInput().setFieldName("nom_copie").setTitle(t('COPY_NAME_LABEL')).setValue(t('COPY_OF', item.title)))
    .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.RADIO_BUTTON).setTitle(t('DESTINATION_TITLE')).setFieldName("emplacement_dest")
      .addItem(t('DEST_SAME'), "meme", true).addItem(t('DEST_ROOT'), "racine", false).addItem(t('DEST_CUSTOM'), "personnalise", false))
    .addWidget(CardService.newTextInput().setFieldName("id_dest_personnalise").setTitle(t('DEST_CUSTOM_INT')).setHint(t('DEST_CUSTOM_HINT')))
    .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.CHECK_BOX).setFieldName("copier_fichiers").addItem(t('COPY_FILES_OPT'), "oui", false))
    .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.CHECK_BOX).setFieldName("arriere_plan").addItem(t('RUN_BG_OPT'), "oui", true))
    .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.CHECK_BOX).setFieldName("conserver_droits").addItem(t('SYNC_PERM_OPT'), "oui", false)));

  // Advanced section
  carte.addSection(CardService.newCardSection().setHeader(t('ADVANCED_OPTS')).setCollapsible(true).setNumUncollapsibleWidgets(0)
    .addWidget(CardService.newTextInput().setFieldName("exclusions").setTitle(t('EXCLUSIONS_LABEL')).setHint(t('EXCLUSIONS_HINT')))
    .addWidget(CardService.newSelectionInput().setType(CardService.SelectionInputType.CHECK_BOX).setFieldName("utiliser_regex").addItem(t('USE_REGEX_OPT'), "oui", false)));

  // Actions
  const sectionActions = CardService.newCardSection().addWidget(CardService.newTextButton()
    .setText(t('START_BTN')).setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction().setFunctionName("lancerDuplication").setParameters({ idDossier: item.id, nomDossier: item.title })));

  if (PersistanceManager.chargerEtat()) {
    sectionActions.addWidget(CardService.newDivider());
    sectionActions.addWidget(CardService.newTextButton().setText(t('RESUME_BTN')).setOnClickAction(CardService.newAction().setFunctionName("reprendreDuplication")));
  }

  carte.addSection(sectionActions);
  carte.addSection(construireSectionHistorique());
  carte.addSection(construirePiedDePage());
  return carte.build();
}

function apercuDossier(e) {
  const { idDossier, nomDossier } = e.parameters;
  if (!MoteurDuplication.estIdValide(idDossier)) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(construireCarteErreur(t('INVALID_FOLDER_ID')))).build();
  }

  try {
    const stats = { dossiers: 0, fichiers: 0, profondeur: 0, limiteAtteinte: false };
    MoteurDuplication.compterContenuDossier(idDossier, stats, 0);

    const estVolumineux = stats.dossiers > 50 || stats.fichiers > 200 || stats.limiteAtteinte;
    let analyse = estVolumineux 
      ? `<font color="${CONFIG.COULEURS.AVERTISSEMENT}"><b>${t('LARGE_TREE')}</b></font><br>${t('MULTIPLE_PASSES')}`
      : `<font color="${CONFIG.COULEURS.SUCCES}">${t('REASONABLE_SIZE')}</font>`;

    if (stats.limiteAtteinte) {
      analyse += `<br><br><font color="${CONFIG.COULEURS.SUR_SURFACE_VARIANTE}"><i>${t('LIMITED_PREVIEW', CONFIG.SEUIL_LIMITE_APERCU.toString())}</i></font>`;
    }

    const carte = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle(t('PREVIEW_BTN')).setSubtitle(nomDossier).setImageUrl(CONFIG.ICONES.LOGO).setImageStyle(CardService.ImageStyle.CIRCLE))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newDecoratedText().setTopLabel(t('SUBFOLDERS')).setText(`${stats.dossiers}${stats.limiteAtteinte ? "+" : ""}`))
        .addWidget(CardService.newDecoratedText().setTopLabel(t('FILES')).setText(`${stats.fichiers}${stats.limiteAtteinte ? "+" : ""}`))
        .addWidget(CardService.newDecoratedText().setTopLabel(t('DEPTH')).setText(t('LEVELS', stats.profondeur.toString())))
        .addWidget(CardService.newDivider())
        .addWidget(CardService.newTextParagraph().setText(analyse))
        .addWidget(CardService.newTextButton().setText(t('BACK_BTN')).setOnClickAction(CardService.newAction().setFunctionName("retourArriere"))))
      .addSection(construirePiedDePage())
      .build();

    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(carte)).build();
  } catch (err) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(construireCarteErreur(t('UNABLE_TO_READ_FOLDER', err.message)))).build();
  }
}

function retourArriere() {
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popCard()).build();
}

function lancerDuplication(e) {
  PersistanceManager.supprimerEtat();
  const { idDossier, nomDossier } = e.parameters;
  const inputs = e.formInputs || {};

  const getVal = (field, fallback = "") => (inputs[field] && inputs[field][0]) ? inputs[field][0] : fallback;
  const isChecked = (field) => inputs[field] && inputs[field][0] === "oui";

  const nomDestination = getVal("nom_copie", t('COPY_OF', nomDossier)).trim();
  const copierFichiers = isChecked("copier_fichiers");
  const executerArriereplan = isChecked("arriere_plan");
  const conserverDroits = isChecked("conserver_droits");
  const emplacementDest = getVal("emplacement_dest", "meme");
  const idDestPersonnalise = getVal("id_dest_personnalise").trim();
  const exclusionsBrutes = getVal("exclusions");
  const utiliserRegex = isChecked("utiliser_regex");
  const exclusions = exclusionsBrutes.split(",").map(s => s.trim()).filter(Boolean);

  if (!nomDestination) return construireCarteErreur(t('EMPTY_COPY_NAME'));
  if (!MoteurDuplication.estIdValide(idDossier)) return construireCarteErreur(t('INVALID_SOURCE_ID'));

  try {
    let idDossierParent = 'root';
    let dossierSourceInfo;

    try {
      dossierSourceInfo = MoteurDuplication.avecRetentative(() => Drive.Files.get(idDossier, { fields: 'parents', supportsAllDrives: true }));
    } catch (err) {
      return construireCarteErreur(t('INVALID_SOURCE_ID'));
    }

    if (emplacementDest === "personnalise" && idDestPersonnalise) {
      if (!MoteurDuplication.estIdValide(idDestPersonnalise)) return construireCarteErreur(t('INVALID_DEST_CHARS'));
      try {
        MoteurDuplication.avecRetentative(() => Drive.Files.get(idDestPersonnalise, { fields: 'id', supportsAllDrives: true }));
        idDossierParent = idDestPersonnalise;
      } catch (err) {
        return construireCarteErreur(t('INVALID_DEST_ID'));
      }
    } else if (emplacementDest === "meme") {
      idDossierParent = (dossierSourceInfo.parents && dossierSourceInfo.parents.length > 0) ? dossierSourceInfo.parents[0] : 'root';
    }

    const dossierDestination = MoteurDuplication.avecRetentative(() => Drive.Files.create({
      name: nomDestination,
      mimeType: CONFIG.MIME_DOSSIER,
      parents: [idDossierParent]
    }, null, { supportsAllDrives: true }));

    if (conserverDroits) MoteurDuplication.synchroniserDroits(idDossier, dossierDestination.id);

    /** @type {DuplicationTask} */
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
      statistiques: { dossiers: 0, fichiers: 0, erreurs: [], ignores: 0 },
      fileAttente: [[idDossier, dossierDestination.id]],
      indexFile: 0
    };

    const resultat = MoteurDuplication.executer(tache, new Date().getTime());
    if (resultat.termine) {
      PersistanceManager.supprimerEtat();
      sauvegarderHistorique(nomDossier, nomDestination, dossierDestination.id, resultat.statistiques);
      notifierParEmail(tache.emailUtilisateur, nomDestination, dossierDestination.id, resultat.statistiques, true);
      return construireCarteResultat(nomDestination, dossierDestination.id, resultat.statistiques, copierFichiers);
    } else {
      PersistanceManager.sauvegarderEtat(resultat.tache);
      if (executerArriereplan) {
        PersistanceManager.programmerArriereplan();
        return construireCarteArriereplan(dossierDestination.id);
      } else {
        return construireCarteProgression(resultat.tache);
      }
    }
  } catch (err) {
    PersistanceManager.supprimerEtat();
    return construireCarteErreur(err.message);
  }
}

function reprendreDuplication() {
  const tache = PersistanceManager.chargerEtat();
  if (!tache) return construireCarteErreur(t('NO_DUPLICATION_CORRUPTED'));

  const resultat = MoteurDuplication.executer(tache, new Date().getTime());
  if (resultat.termine) {
    PersistanceManager.supprimerEtat();
    sauvegarderHistorique(tache.nomSource, tache.nomDestination, tache.idDossierDest, resultat.statistiques);
    notifierParEmail(tache.emailUtilisateur, tache.nomDestination, tache.idDossierDest, resultat.statistiques, true);
    return construireCarteResultat(tache.nomDestination, tache.idDossierDest, resultat.statistiques, tache.copierFichiers);
  } else {
    PersistanceManager.sauvegarderEtat(resultat.tache);
    if (tache.executerArriereplan) {
      PersistanceManager.programmerArriereplan();
      return construireCarteArriereplan(tache.idDossierDest);
    }
    return construireCarteProgression(resultat.tache);
  }
}

function traiterTacheArriereplan() {
  PersistanceManager.nettoyerTriggers();
  const tache = PersistanceManager.chargerEtat();
  if (!tache) return;

  try {
    const resultat = MoteurDuplication.executer(tache, new Date().getTime(), true);
    if (resultat.termine) {
      PersistanceManager.supprimerEtat();
      sauvegarderHistorique(tache.nomSource, tache.nomDestination, tache.idDossierDest, resultat.statistiques);
      notifierParEmail(tache.emailUtilisateur, tache.nomDestination, tache.idDossierDest, resultat.statistiques, true);
    } else {
      PersistanceManager.sauvegarderEtat(resultat.tache);
      PersistanceManager.programmerArriereplan();
    }
  } catch (e) {
    notifierParEmail(tache.emailUtilisateur, tache.nomDestination, tache.idDossierDest, tache.statistiques, false, e.message);
  }
}

// UI Elements Construction
function construirePiedDePage() {
  return CardService.newCardSection().setCollapsible(false)
    .addWidget(CardService.newTextParagraph().setText(`<font color="${CONFIG.COULEURS.SUR_SURFACE_VARIANTE}"><i>${CONFIG.VERSION} · ${CONFIG.AUTEUR}</i></font>`));
}

function construireCarteArriereplan(idDossierDest) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(t('BG_PROCESSING_TITLE')).setSubtitle(t('BG_PROCESSING_SUB')).setImageUrl(CONFIG.ICONES.LOGO).setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(`<font color="${CONFIG.COULEURS.PRIMAIRE}"><b>${t('BG_PROCESSING_MSG1')}</b></font><br><br>${t('BG_PROCESSING_MSG2')}`))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextButton().setText(t('OPEN_FOLDER_BUILD')).setOpenLink(CardService.newOpenLink().setUrl(`https://drive.google.com/drive/folders/${idDossierDest}`))))
    .addSection(construirePiedDePage()).build();
}

function construireCarteProgression(tache) {
  const reste = tache.fileAttente.length - (tache.indexFile || 0);
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(t('PASS_INCOMPLETE_TITLE')).setSubtitle(t('PASS_INCOMPLETE_SUB')).setImageUrl(CONFIG.ICONES.LOGO).setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText().setTopLabel(t('FOLDERS_CREATED')).setText(`${tache.statistiques.dossiers}`))
      .addWidget(CardService.newDecoratedText().setTopLabel(t('FILES_COPIED')).setText(`${tache.statistiques.fichiers || 0}`))
      .addWidget(CardService.newDecoratedText().setTopLabel(t('REMAINING_TO_PROCESS')).setText(t('FOLDERS_IN_QUEUE', reste.toString())))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextParagraph().setText(`<font color="${CONFIG.COULEURS.AVERTISSEMENT}">${t('LARGE_TREE_CONTINUE')}</font>`))
      .addWidget(CardService.newTextButton().setText(t('CONTINUE_BTN')).setTextButtonStyle(CardService.TextButtonStyle.FILLED).setOnClickAction(CardService.newAction().setFunctionName("reprendreDuplication"))))
    .addSection(construirePiedDePage()).build();
}

function construireCarteResultat(nomDest, idDest, statistiques, copierFichiers) {
  const section = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText().setTopLabel(t('FOLDER_CREATED')).setText(nomDest).setWrapText(true))
    .addWidget(CardService.newDecoratedText().setTopLabel(t('SUBFOLDERS_CREATED')).setText(`${statistiques.dossiers}`));
    
  if (copierFichiers) section.addWidget(CardService.newDecoratedText().setTopLabel(t('FILES_COPIED')).setText(`${statistiques.fichiers || 0}`));
  if (statistiques.ignores > 0) section.addWidget(CardService.newDecoratedText().setTopLabel(t('FOLDERS_IGNORED')).setText(`${statistiques.ignores}`));
  
  if (statistiques.erreurs && statistiques.erreurs.length > 0) {
    let errTxt = `<font color="${CONFIG.COULEURS.ERREUR}"><b>${t('ERRORS_ENCOUNTERED', statistiques.erreurs.length.toString())}</b></font><br>`;
    statistiques.erreurs.slice(0, 5).forEach(err => { errTxt += `• <i>${err}</i><br>`; });
    if (statistiques.erreurs.length > 5) errTxt += `<i>${t('AND_OTHERS', (statistiques.erreurs.length - 5).toString())}</i>`;
    section.addWidget(CardService.newTextParagraph().setText(errTxt));
  }

  section.addWidget(CardService.newDivider());
  section.addWidget(CardService.newTextButton().setText(t('OPEN_CREATED_FOLDER')).setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOpenLink(CardService.newOpenLink().setUrl(`https://drive.google.com/drive/folders/${idDest}`)));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(t('DUPLICATION_FINISHED_TITLE')).setSubtitle(t('FOLDERS_CREATED_SUCCESS', statistiques.dossiers.toString())).setImageUrl(CONFIG.ICONES.LOGO).setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(section).addSection(construirePiedDePage()).build();
}

function construireCarteErreur(message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(t('ERROR_TITLE')).setSubtitle(t('PROBLEM_OCCURRED')).setImageUrl(CONFIG.ICONES.LOGO).setImageStyle(CardService.ImageStyle.CIRCLE))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(`<font color="${CONFIG.COULEURS.ERREUR}">${message}</font>`))
      .addWidget(CardService.newDivider())
      .addWidget(CardService.newTextButton().setText(t('BACK_BTN')).setOnClickAction(CardService.newAction().setFunctionName("retourArriere"))))
    .addSection(construirePiedDePage()).build();
}

function construireSectionHistorique() {
  const section = CardService.newCardSection().setHeader(t('HISTORY_HEADER')).setCollapsible(true).setNumUncollapsibleWidgets(0);
  const historique = chargerHistoriqueRaw();

  if (historique.length === 0) {
    return section.addWidget(CardService.newTextParagraph().setText(`<font color="${CONFIG.COULEURS.SUR_SURFACE_VARIANTE}">${t('HISTORY_EMPTY')}</font>`));
  }

  historique.slice(0, CONFIG.MAX_ENTREES_HISTORIQUE).forEach(e => {
    section.addWidget(CardService.newDecoratedText()
      .setTopLabel(e.date).setText(`${e.nomSource} → ${e.nomDestination}`)
      .setBottomLabel(`${e.dossiers} ${t('FOLDERS_CREATED').toLowerCase()}, ${e.fichiers} ${t('FILES_COPIED').toLowerCase()}`).setWrapText(true)
      .setButton(CardService.newImageButton().setIconUrl(CONFIG.ICONES.OUVRIR_NOUVEAU).setAltText(t('HISTORY_OPEN_ALT')).setOpenLink(CardService.newOpenLink().setUrl(`https://drive.google.com/drive/folders/${e.idDest}`))));
  });

  section.addWidget(CardService.newDivider());
  section.addWidget(CardService.newTextButton().setText(t('CLEAR_HISTORY_BTN')).setOnClickAction(CardService.newAction().setFunctionName("effacerHistorique")));
  return section;
}

function sauvegarderHistorique(nomSource, nomDestination, idDest, statistiques) {
  const historique = chargerHistoriqueRaw();
  historique.unshift({
    date: new Date().toLocaleString(Session.getActiveUserLocale() || "en-US"),
    nomSource, nomDestination, idDest, dossiers: statistiques.dossiers, fichiers: statistiques.fichiers || 0
  });
  PersistanceManager.getStore().setProperty("historique", JSON.stringify(historique.slice(0, CONFIG.MAX_ENTREES_HISTORIQUE)));
}

function chargerHistoriqueRaw() {
  const brut = PersistanceManager.getStore().getProperty("historique");
  if (!brut) return [];
  try { return JSON.parse(brut); } catch (e) { return []; }
}

function effacerHistorique() {
  PersistanceManager.getStore().deleteProperty("historique");
  return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText(t('HISTORY_CLEARED'))).setStateChanged(true).build();
}

// Mail Notification Engine
function notifierParEmail(email, nomDest, idDest, statistiques, estSucces, detailErreur = "") {
  if (!email) return;
  const url = `https://drive.google.com/drive/folders/${idDest}`;
  const subject = estSucces ? t('EMAIL_SUCCESS_SUBJECT', nomDest) : t('EMAIL_FAILED_SUBJECT', nomDest);
  
  let corpsTexte = `${t('EMAIL_SUCCESS_HELLO')},\n\n`;
  if (estSucces) {
    corpsTexte += `${t('EMAIL_SUCCESS_BODY_1')}\n\n${t('EMAIL_SUCCESS_FOLDER', nomDest)}\n${t('FOLDERS_CREATED')} : ${statistiques.dossiers}\n${t('FILES_COPIED')} : ${statistiques.fichiers || 0}\n\n${url}`;
  } else {
    corpsTexte += `${t('EMAIL_FAILED_BODY_1')}\n\n${detailErreur}\n\n${t('EMAIL_FAILED_BODY_2')}`;
  }

  // Génération simplifiée et robuste du HTML adaptatif (Material Blue / Red Theme)
  const codeCouleur = estSucces ? "#0b57d0" : "#b3261e";
  const iconeHeader = estSucces ? "&#9989;" : "&#9888;";
  const titreHeader = estSucces ? t('DUPLICATION_FINISHED_TITLE') : t('EMAIL_FAILED_TITLE');
  const sousTitreHeader = estSucces ? t('EMAIL_SUCCESS_SUBTITLE') : nomDest;

  const htmlBody = `
  <!DOCTYPE html>
  <html>
  <body style="margin:0;padding:0;background-color:#f3f6fc;font-family:Roboto,Arial,sans-serif;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#f3f6fc;padding:24px 0;">
      <tr><td align="center">
        <table width="560" cellpadding="0" cellspacing="0" style="background-color:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.12);">
          <tr>
            <td style="background-color:${codeCouleur};padding:32px 24px;text-align:center;color:#FFFFFF;">
              <div style="font-size:28px;">${iconeHeader}</div>
              <h1 style="margin:8px 0 0;font-size:22px;font-weight:500;">${titreHeader}</h1>
              <p style="margin:4px 0 0;font-size:14px;color:rgba(255,255,255,0.85);">${sousTitreHeader}</p>
            </td>
          </tr>
          <tr>
            <td style="padding:24px;color:#202124;font-size:14px;line-height:1.6;">
              <p>${t('EMAIL_SUCCESS_HELLO')},</p>
              <p>${estSucces ? t('EMAIL_SUCCESS_BODY_2') : t('EMAIL_FAILED_BODY_1_HTML')}</p>
              
              ${estSucces ? `
              <table width="100%" style="background-color:#f3f6fc;border-radius:12px;margin:20px 0;padding:16px;">
                <tr><td><b>${t('FOLDER_CREATED')} :</b> ${nomDest}</td></tr>
                <tr><td><b>${t('SUBFOLDERS')} :</b> ${statistiques.dossiers}</td></tr>
                <tr><td><b>${t('FILES')} :</b> ${statistiques.fichiers || 0}</td></tr>
              </table>
              <div style="text-align:center;margin:24px 0;">
                <a href="${url}" style="padding:12px 32px;background-color:#0b57d0;color:#FFFFFF;text-decoration:none;border-radius:100px;font-weight:500;">${t('EMAIL_SUCCESS_OPEN_BTN')}</a>
              </div>
              ` : `
              <div style="background-color:#FEF7F6;border:1px solid #F5C6CB;border-radius:12px;padding:16px;color:#b3261e;font-family:monospace;margin:20px 0;">
                <b>${t('EMAIL_FAILED_DETAIL')} :</b><br>${detailErreur}
              </div>
              <table width="100%" style="background-color:#f3f6fc;border-radius:12px;padding:16px;margin:20px 0;">
                <tr><td><b>${t('EMAIL_FAILED_ACTION_TITLE')} :</b><br>${t('EMAIL_FAILED_ACTION_DESC')}</td></tr>
              </table>
              `}
            </td>
          </tr>
          <tr>
            <td style="padding:16px;background-color:#f3f6fc;text-align:center;font-size:12px;color:#9AA0A6;border-top:1px solid #E8EAED;">
              ${CONFIG.VERSION} · ${CONFIG.AUTEUR}<br>${t('EMAIL_FOOTER_DESC')}
            </td>
          </tr>
        </table>
      </td></tr>
    </table>
  </body>
  </html>`;

  try {
    MailApp.sendEmail({ to: email, subject, body: corpsTexte, htmlBody: htmlBody });
  } catch (e) {
    console.error("Échec d'envoi de la notification courriel: " + e.message);
  }
}

function autoriserMaintenant() {
  const racine = Drive.Files.get('root', { fields: 'name' });
  console.log("Autorisation validée sur le Drive racine : " + racine.name);
}

