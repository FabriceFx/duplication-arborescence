/**
 * ============================================================================
 * MODULE 3 : MOTEUR ALGORITHMIQUE DE PARCOURS ET DUPLICATION
 * ============================================================================
 */
const MoteurDuplication = {

  /**
   * Exécute les retentatives automatiques sur les appels de services avancés.
   */
  avecRetentative(action, maxTentatives = 3) {
    for (let tentative = 1; tentative <= maxTentatives; tentative++) {
      try {
        return action();
      } catch (e) {
        if (tentative >= maxTentatives) throw e;
        const delaiMs = Math.pow(2, tentative) * 1000 + Math.round(Math.random() * 500);
        Utilities.sleep(delaiMs);
      }
    }
  },

  estIdValide(id) {
    return typeof id === "string" && id.length > 0 && CONFIG.REGEX_VALIDATION_ID.test(id);
  },

  /**
   * Analyse prédictive de taille de l'arborescence (Récursion sécurisée par seuil).
   */
  compterContenuDossier(idDossier, statistiques, profondeur) {
    if ((statistiques.dossiers + statistiques.fichiers) > CONFIG.SEUIL_LIMITE_APERCU || statistiques.limiteAtteinte) {
      statistiques.limiteAtteinte = true;
      return;
    }
    if (profondeur > statistiques.profondeur) statistiques.profondeur = profondeur;

    let token = null;
    const sousDossiers = [];

    do {
      const reponse = this.avecRetentative(() => Drive.Files.list({
        q: `'${idDossier}' in parents and trashed = false`,
        fields: "nextPageToken, files(id, mimeType)",
        pageSize: 1000,
        pageToken: token,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      }));

      const files = reponse.files || [];
      for (const file of files) {
        if (file.mimeType === CONFIG.MIME_DOSSIER) {
          statistiques.dossiers++;
          sousDossiers.push(file.id);
        } else if (file.mimeType !== CONFIG.MIME_RACCOURCI) {
          statistiques.fichiers++;
        }
      }
      if ((statistiques.dossiers + statistiques.fichiers) > CONFIG.SEUIL_LIMITE_APERCU) {
        statistiques.limiteAtteinte = true;
      }
      token = reponse.nextPageToken;
    } while (token && !statistiques.limiteAtteinte);

    for (const idSub of sousDossiers) {
      if (statistiques.limiteAtteinte) break;
      this.compterContenuDossier(idSub, statistiques, profondeur + 1);
    }
  },

  /**
   * Algorithme BFS optimisé avec pointeur d'index (Complexité temporelle linéaire O(N)).
   * @param {DuplicationTask} tache
   * @param {number} heureDebut - Timestamp de départ.
   * @param {boolean} estArriereplan
   * @return {{termine: boolean, tache: DuplicationTask, statistiques: TaskStats}}
   */
  executer(tache, heureDebut, estArriereplan = false) {
    const queue = tache.fileAttente;
    let index = tache.indexFile || 0; // Utilisation d'un pointeur au lieu de .shift()
    const dureeMax = estArriereplan ? CONFIG.DUREE_MAX_MS_ARRPLAN : CONFIG.DUREE_MAX_MS_UI;

    // Compilation initiale des règles de filtrage
    const filtres = (tache.exclusions || []).map(pattern => {
      if (tache.utiliserRegex) {
        try { return new RegExp(pattern, "i"); } catch(e) { return null; }
      }
      return pattern.toLowerCase().trim();
    }).filter(Boolean);

    const estExclu = (nom) => {
      if (filtres.length === 0) return false;
      return filtres.some(f => f instanceof RegExp ? f.test(nom) : nom.toLowerCase().trim() === f);
    };

    const collecterErreur = (msg) => {
      console.error("Échec Moteur: " + msg);
      if (tache.statistiques.erreurs.length < CONFIG.MAX_ERREURS_STOCKEES) {
        tache.statistiques.erreurs.push(msg.length > 100 ? `${msg.substring(0, 97)}...` : msg);
      }
    };

    while (index < queue.length) {
      // Vérification stricte du timeout GAS
      if ((new Date().getTime() - heureDebut) > dureeMax) {
        tache.indexFile = index; // On garde en mémoire la position du pointeur
        tache.fileAttente = queue;
        return { termine: false, tache };
      }

      const [idSourceActuel, idDestActuel] = queue[index];
      index++; // On avance dans la file
      let token = null;

      do {
        try {
          const reponse = this.avecRetentative(() => Drive.Files.list({
            q: `'${idSourceActuel}' in parents and trashed = false`,
            fields: "nextPageToken, files(id, name, mimeType, description, folderColorRgb)",
            pageSize: 1000,
            pageToken: token,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true
          }));

          const enfants = reponse.files || [];
          for (const enfant of enfants) {
            if (enfant.mimeType === CONFIG.MIME_DOSSIER) {
              if (estExclu(enfant.name)) {
                tache.statistiques.ignores++;
                continue;
              }
              try {
                const meta = { name: enfant.name, mimeType: CONFIG.MIME_DOSSIER, parents: [idDestActuel] };
                if (enfant.description) meta.description = enfant.description;
                if (enfant.folderColorRgb) meta.folderColorRgb = enfant.folderColorRgb;

                const nouveauDossier = this.avecRetentative(() => Drive.Files.create(meta, null, { supportsAllDrives: true }));
                if (tache.conserverDroits) this.synchroniserDroits(enfant.id, nouveauDossier.id);

                tache.statistiques.dossiers++;
                queue.push([enfant.id, nouveauDossier.id]);
              } catch (errDossier) {
                collecterErreur(`Dossier "${enfant.name}": ${errDossier.message}`);
              }

            } else if (enfant.mimeType !== CONFIG.MIME_RACCOURCI && tache.copierFichiers) {
              try {
                const metaFichier = { parents: [idDestActuel], name: enfant.name };
                if (enfant.description) metaFichier.description = enfant.description;

                const nouveauFichier = this.avecRetentative(() => Drive.Files.copy(metaFichier, enfant.id, { supportsAllDrives: true }));
                if (tache.conserverDroits) this.synchroniserDroits(enfant.id, nouveauFichier.id);

                tache.statistiques.fichiers++;
              } catch (errFichier) {
                collecterErreur(`Fichier "${enfant.name}": ${errFichier.message}`);
              }
            }
          }
          token = reponse.nextPageToken;
        } catch (e) {
          collecterErreur("Erreur critique d'accès au dossier parent: " + e.message);
          break;
        }
      } while (token);
    }

    tache.indexFile = index;
    return { termine: true, statistiques: tache.statistiques };
  },

  synchroniserDroits(idSource, idDest) {
    try {
      const liste = this.avecRetentative(() => Drive.Permissions.list(idSource, {
        supportsAllDrives: true,
        fields: "permissions(emailAddress, role, type, deleted)"
      }));
      const permissions = liste.permissions || [];
      
      permissions.forEach(p => {
        if (p.deleted || !p.emailAddress || p.role === 'owner') return;
        try {
          this.avecRetentative(() => Drive.Permissions.create({
            role: p.role,
            type: p.type,
            emailAddress: p.emailAddress
          }, idDest, { supportsAllDrives: true, sendNotificationEmail: false }));
        } catch (err) {
          console.warn(`ACL non synchronisée pour ${p.emailAddress}: ${err.message}`);
        }
      });
    } catch (e) {
      console.warn(`Échec de lecture des ACL source: ${e.message}`);
    }
  }
};

