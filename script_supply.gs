function executerGrandNettoyageTP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var feuilleSource = ss.getSheetByName("DATA_BRUTE"); 
  var feuilleClean = ss.getSheetByName("DATA_CLEAN-02");
  var feuilleAnalyse = ss.getSheetByName("ANALYSE");
  
  var lastRow = feuilleSource.getLastRow();
  var donnees = feuilleSource.getRange("A2:J" + lastRow).getValues();
  var resultats = [];
  
  // Variables pour les totaux (Étape 4)
  var totalExpeditionsValides = 0;
  var coutTotalFCFA = 0;
  var poidsTotalTonnes = 0;

  for (var i = 0; i < donnees.length; i++) {
    var ligne = donnees[i];

    // 1. Tracking_ID (Col A) -> Format TRK-XXX
    var trk = ligne[0].toString().toUpperCase();
    if (trk.includes("TRK") && !trk.includes("-")) {
      trk = trk.replace("TRK", "TRK-");
    }

    // 2. Villes (Col D : Ville Arrivée) -> Nettoyage + Lomé
    var villeArr = ligne[3].toString().trim();
    villeArr = villeArr.charAt(0).toUpperCase() + villeArr.slice(1).toLowerCase();
    if (villeArr.toLowerCase().includes("lom")) villeArr = "Lomé";

    // 3. Dates & Contrôle Qualité (Col B: Commande, Col I: Arrivée Prévue)
    var dateCmd = new Date(ligne[1]);
    var dateArrPrevue = new Date(ligne[8]);
    var estValide = (dateArrPrevue >= dateCmd);

    // 4. Poids Brut (Col F) -> Conversion numérique en kg
    var poidsRaw = ligne[5].toString().toLowerCase();
    var poidsKg = parseFloat(poidsRaw.replace(",", ".")); // Remplace virgule par point
    if (poidsRaw.includes("t") || poidsRaw.includes("tonne")) {
      poidsKg = poidsKg * 1000;
    }

    // 5. Coûts (Col H : Cout_Transport) -> Nettoyage symboles
    var coutRaw = ligne[7].toString().replace("$", "").replace("USD", "").replace(",", ".").trim();
    var coutUSD = coutRaw === "" || coutRaw === "NULL" ? 0 : parseFloat(coutRaw);

    // --- CALCULS ÉTAPE 3 ---
    var coutFCFA = coutUSD * 600;
    var delai = estValide ? Math.round((dateArrPrevue - dateCmd) / (1000 * 60 * 60 * 24)) : "ERREUR";
    var volume = parseFloat(ligne[6]) || 1; // Col G: Volume_m3
    var ratioDensite = poidsKg / volume;

    // --- COMPTEURS POUR L'ANALYSE ---
    if (estValide && trk !== "") {
      totalExpeditionsValides++;
      coutTotalFCFA += coutFCFA;
      poidsTotalTonnes += (poidsKg / 1000);
    }

    // On prépare la ligne pour DATA_CLEAN
    resultats.push([
      trk,             // A: Tracking_ID
      villeArr,        // B: Ville_Arrivée
      dateCmd,         // C: Date_Commande
      dateArrPrevue,   // D: Date_Arrivée_Prévue
      poidsKg,         // E: Poids (kg)
      coutUSD,         // F: Coût (USD)
      coutFCFA,        // G: Coût (FCFA)
      delai,           // H: Délai (Jours)
      ratioDensite,    // I: Ratio Densité
      estValide ? "VALIDE" : "ILLOGIQUE" // J: Statut
    ]);
  }

  // VIDAGE ET REMPLISSAGE DE DATA_CLEAN
  feuilleClean.clear();
  var entetes = [["Tracking_ID", "Ville Arrivée", "Date Commande", "Date Prévue", "Poids (kg)", "Coût (USD)", "Coût (FCFA)", "Délai (Jours)", "Ratio Densité", "Contrôle Qualité"]];
  feuilleClean.getRange(1, 1, 1, 10).setValues(entetes).setBackground("#444444").setFontColor("white").setFontWeight("bold");
  feuilleClean.getRange(2, 1, resultats.length, 10).setValues(resultats);
  
  // Formatage Dates et Couleurs d'alerte
  feuilleClean.getRange("C2:D").setNumberFormat("dd/mm/yyyy");
  for (var j = 0; j < resultats.length; j++) {
    if (resultats[j][9] === "ILLOGIQUE") {
      feuilleClean.getRange(j + 2, 1, 1, 10).setBackground("#f4cccc").setFontColor("red");
    }
  }

  // MISE À JOUR AUTOMATIQUE DE L'ONGLET ANALYSE (Cellules B2, B3, B4)
  feuilleAnalyse.getRange("B2").setValue(totalExpeditionsValides);
  feuilleAnalyse.getRange("B3").setValue(coutTotalFCFA);
  feuilleAnalyse.getRange("B4").setValue(poidsTotalTonnes);

  SpreadsheetApp.getUi().alert("🚀 Traitement Supply Chain terminé avec succès !");
}
