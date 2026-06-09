/**
 * ============================================================================
 * MODULE 2 : PERSISTANCE ET TRIGGERS D'ARRIÈRE-PLAN
 * ============================================================================
 */
const PersistanceManager = {
  getStore() {
    return PropertiesService.getUserProperties();
  },

  /**
   * Sauvegarde l'état d'un job de duplication par blocs segmentés.
   * @param {DuplicationTask} tache
   */
  sauvegarderEtat(tache) {
    this.supprimerEtat();
    const store = this.getStore();
    const json = JSON.stringify(tache);
    const chunksCount = Math.ceil(json.length / CONFIG.TAILLE_SEGMENT_PROPRIETE);
    
    store.setProperty("nombre_segments_tache", chunksCount.toString());
    for (let i = 0; i < chunksCount; i++) {
      const debut = i * CONFIG.TAILLE_SEGMENT_PROPRIETE;
      store.setProperty(`segment_tache_${i}`, json.substring(debut, debut + CONFIG.TAILLE_SEGMENT_PROPRIETE));
    }
  },

  /**
   * Charge et reconstitue l'état complet du job.
   * @return {DuplicationTask|null}
   */
  chargerEtat() {
    const store = this.getStore();
    const countStr = store.getProperty("nombre_segments_tache");
    if (!countStr) return null;
    
    const count = parseInt(countStr, 10);
    let json = "";
    for (let i = 0; i < count; i++) {
      json += store.getProperty(`segment_tache_${i}`) || "";
    }
    
    try {
      return JSON.parse(json);
    } catch (e) {
      console.error("État corrompu détruit.", e);
      this.supprimerEtat();
      return null;
    }
  },

  supprimerEtat() {
    const store = this.getStore();
    const countStr = store.getProperty("nombre_segments_tache");
    if (countStr) {
      const count = parseInt(countStr, 10);
      for (let i = 0; i < count; i++) {
        store.deleteProperty(`segment_tache_${i}`);
      }
      store.deleteProperty("nombre_segments_tache");
    }
  },

  /**
   * Nettoie les déclencheurs orphelins liés à l'arrière-plan.
   */
  nettoyerTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
      if (t.getHandlerFunction() === "traiterTacheArriereplan") {
        try {
          ScriptApp.deleteTrigger(t);
        } catch (e) {
          console.warn("Échec suppression trigger : " + e.message);
        }
      }
    }
  },

  programmerArriereplan() {
    this.nettoyerTriggers();
    try {
      ScriptApp.newTrigger("traiterTacheArriereplan")
        .timeBased()
        .after(60 * 1000)
        .create();
    } catch (e) {
      console.error("Erreur fatale planification trigger : " + e.message);
    }
  }
};

