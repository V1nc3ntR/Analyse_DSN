// DateUtils.gs - Utilitaires de gestion des dates

const DateUtils = (function() {
  return {
    /**
     * Obtient tous les mois entre deux dates
     * @param {Date} dateDebut - Date de début
     * @param {Date} dateFin - Date de fin
     * @return {array} - Liste des mois au format {annee, mois}
     */
    getMonthsBetween: function(dateDebut, dateFin) {
      var months = [];
      var start = new Date(dateDebut.getFullYear(), dateDebut.getMonth(), 1);
      var end   = new Date(dateFin.getFullYear(), dateFin.getMonth(), 1);
      while(start <= end) {
        months.push({
          annee: start.getFullYear(),
          mois : start.getMonth() + 1
        });
        start.setMonth(start.getMonth() + 1);
      }
      return months;
    },
    
    /**
     * Parse une chaîne de date dans différents formats
     * @param {string} str - Chaîne représentant une date
     * @return {Date|null} - Objet Date ou null si invalide
     */
    parseDateString: function(str) {
      if (!str) return null;
      
      // Format JJMMAAAA (exemple: 01022020)
      if(/^\d{8}$/.test(str)) {
        var jj = str.substring(0,2);
        var mm = str.substring(2,4);
        var yyyy = str.substring(4,8);
        var d = new Date(Number(yyyy), Number(mm)-1, Number(jj));
        if(!isNaN(d.getTime())) return d;
      }
      
      // Format ISO (exemple: 2020-02-01)
      var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if(iso) {
        var d2 = new Date(Number(iso[1]), Number(iso[2])-1, Number(iso[3]));
        if(!isNaN(d2.getTime())) return d2;
      }
      
      // Format FR (exemple: 01/02/2020)
      var fr = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if(fr) {
        var d3 = new Date(Number(fr[3]), Number(fr[2])-1, Number(fr[1]));
        if(!isNaN(d3.getTime())) return d3;
      }
      
      // Dernière tentative avec le constructeur Date
      var dd = new Date(str);
      return isNaN(dd.getTime()) ? null : dd;
    },
    
    /**
     * Obtient le dernier jour d'un mois
     * @param {string} moisPrincipal - Mois au format "YYYY-MM"
     * @return {Date} - Date du dernier jour
     */
    getLastDayOfMonth: function(moisPrincipal) {
      var parts = moisPrincipal.split("-");
      if(parts.length < 2) return new Date();
      var yyyy = Number(parts[0]);
      var mm = Number(parts[1]);
      return new Date(yyyy, mm, 0);
    },
    
    /**
     * Calcule l'âge à partir d'une date de naissance
     * @param {Date} dateNais - Date de naissance
     * @param {Date} refDate - Date de référence
     * @return {number} - Âge calculé
     */
    calcAge: function(dateNais, refDate) {
      var age = refDate.getFullYear() - dateNais.getFullYear();
      var m = refDate.getMonth() - dateNais.getMonth();
      if(m < 0 || (m === 0 && refDate.getDate() < dateNais.getDate())) {
        age--;
      }
      return age;
    },
    
    /**
     * Détermine la tranche d'âge
     * @param {number} age - Âge
     * @return {string} - Tranche d'âge
     */
    getTrancheAge: function(age) {
      if(age < 30) return "moins de 30 ans";
      if(age < 40) return "30 à 39 ans";
      if(age < 50) return "40 à 49 ans";
      return "50 ans et plus";
    }
  };
})();

// Exportation des fonctions pour compatibilité
function parseDateString(str) {
  return DateUtils.parseDateString(str);
}

function getLastDayOfMonth(moisPrincipal) {
  return DateUtils.getLastDayOfMonth(moisPrincipal);
}

function calcAge(dateNais, refDate) {
  return DateUtils.calcAge(dateNais, refDate);
}

function getTrancheAge(age) {
  return DateUtils.getTrancheAge(age);
}

function getMonthsBetween(dateDebut, dateFin) {
  return DateUtils.getMonthsBetween(dateDebut, dateFin);
}
