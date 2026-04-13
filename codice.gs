// @ts-nocheck
function importaEventiCalendarioAvanzato() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foglioDestinazione = ss.getSheetByName("TAM");

  // --- CONFIGURAZIONE ---
  var calId = "b04892389811164a166e6c39cfb4958009b04870d49870ecd80a02ffbbbf1bf0@group.calendar.google.com";
  var urlFileNominativi = "https://docs.google.com/spreadsheets/d/1D29qfRM5BMXiAs7iYlb7gqy1nYttMH-Czmu5WeEfH1c/edit";

  var emailAccesso = Session.getActiveUser().getEmail().toLowerCase().trim();

  // 1. Intervallo temporale: DA OGGI per 2 mesi
  var inizio = new Date();
  inizio.setHours(0, 0, 0, 0);
  var fine = new Date();
  fine.setMonth(inizio.getMonth() + 2);

  var calendario = CalendarApp.getCalendarById(calId);
  if (!calendario) {
    SpreadsheetApp.getUi().alert("Calendario non trovato!");
    return;
  }

  // --- 2. RICERCA NOMINATIVO (Colonna B Nome, Colonna I Email) ---
  var nomeTrovato = "Non trovato";
  try {
    var ssNominativi = SpreadsheetApp.openByUrl(urlFileNominativi);
    var foglioNominativi = ssNominativi.getSheetByName("Nominativi");
    var dati = foglioNominativi.getDataRange().getValues();

    for (var i = 0; i < dati.length; i++) {
      if (dati[i][8]) { // Colonna I
        var emailNelFoglio = dati[i][8].toString().trim().toLowerCase();
        if (emailNelFoglio === emailAccesso) {
          nomeTrovato = dati[i][1]; // Colonna B
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("Errore ricerca nome: " + e.message);
  }

  // --- 3. RECUPERO E SCRITTURA EVENTI ---
  var eventi = calendario.getEvents(inizio, fine);

  // Pulizia foglio TAM
  if (foglioDestinazione.getLastRow() > 1) {
    foglioDestinazione.getRange(2, 1, foglioDestinazione.getLastRow(), 3).clearContent();
  }

  if (eventi.length === 0) return;

  var output = [];
  for (var j = 0; j < eventi.length; j++) {
    var ev = eventi[j];

    // Estrazione orari
    var oraInizio = Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "HH:mm");
    var oraFine = Utilities.formatDate(ev.getEndTime(), Session.getScriptTimeZone(), "HH:mm");

    // Punto Finale: Prefisso "TAM - " + Titolo + Orari inizio-fine
    var descrizioneCompleta = "TAM - " + ev.getTitle() + " " + oraInizio + " - " + oraFine;

    output.push([
      ev.getStartTime(),   // Colonna A: Data
      descrizioneCompleta,  // Colonna B: "TAM - Evento 09:00 - 10:00"
      nomeTrovato           // Colonna C: Nominativo
    ]);
  }

  // Scrittura finale nel foglio
  foglioDestinazione.getRange(2, 1, output.length, 3).setValues(output);

  // Formattazione data pulita (gg/mm/aaaa)
  foglioDestinazione.getRange(2, 1, output.length, 1).setNumberFormat("dd/mm/yyyy");

  // Logger.log("Sincronizzazione completata con prefisso TAM.");
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGruppi = ss.getSheetByName("Gruppi");
  var sheetDB = ss.getSheetByName("Database_PW");
  var azione = e.parameter.azione;

  if (!azione) {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Gestione TePu')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // --- AZIONE: RICERCA ---
  if (azione === "cerca") {
    var input = e.parameter.email.toLowerCase().trim();
    var dataG = sheetGruppi.getDataRange().getValues();
    for (var r = 1; r < dataG.length; r++) {
      if (dataG[r][0].toString().toLowerCase().trim() === input || dataG[r][2].toString().toLowerCase().trim() === input) {
        return successo({
          trovato: true,
          nome: dataG[r][0],
          email: dataG[r][2].toString().trim()
        });
      }
    }
    return successo({ trovato: false });
  }

  // --- AZIONE: LOGIN ---
  if (azione === "login") {
    var identificativo = e.parameter.telefono.trim();
    var passIn = e.parameter.password.trim();
    var dataDB = sheetDB.getDataRange().getValues();

    for (var r = 1; r < dataDB.length; r++) {
      if (dataDB[r][0].toString().trim().toLowerCase() === identificativo.toLowerCase()) {
        if (passIn === dataDB[r][2].toString().trim()) {
          var deveCambiare = (dataDB[r][3] === "" || dataDB[r][3] === undefined);

          // Recupero link e abilitazioni dal Database_PW
          var linkTePuDB = dataDB[r][6] || "";        // Colonna G
          var linkTerritoriDB = dataDB[r][7] || "";   // Colonna H
          // Controllo Colonna I per abilitazione pulsanti TePu
          var abilitatoTePu = (dataDB[r][8] && dataDB[r][8].toString().trim().toLowerCase() === "si");

          var emailUtente = "";
          var linkAppSheetIncarichi = "";
          var dataGruppi = sheetGruppi.getDataRange().getValues();

          for (var i = 1; i < dataGruppi.length; i++) {
            if (dataGruppi[i][0].toString().trim().toLowerCase() === identificativo.toLowerCase()) {
              emailUtente = dataGruppi[i][2] || "";
              linkAppSheetIncarichi = dataGruppi[i][7] || "";
              break;
            }
          }

          return successo({
            login: true,
            reset: deveCambiare,
            email: emailUtente,
            linkAppSheet: linkAppSheetIncarichi,
            linkTePu: linkTePuDB,
            linkTerritori: linkTerritoriDB,
            mostraTePu: abilitatoTePu // Cruciale per far vedere i tasti in index.html
          });
        }
      }
    }
    return successo({ login: false });
  }

  // --- AZIONE: CAMBIA PASSWORD ---
  if (azione === "cambiaPassword") {
    var identificativo = e.parameter.telefono.trim();
    var nuova = e.parameter.nuovaPass.trim();
    var dataDB = sheetDB.getDataRange().getValues();
    for (var r = 1; r < dataDB.length; r++) {
      if (dataDB[r][0].toString().trim().toLowerCase() === identificativo.toLowerCase()) {
        sheetDB.getRange(r + 1, 3).setValue(nuova);
        sheetDB.getRange(r + 1, 4).setValue(nuova);
        return successo({ aggiornato: true });
      }
    }
    return successo({ aggiornato: false });
  }

  // --- NUOVA AZIONE DA AGGIUNGERE SUBITO DOPO ---
  if (azione === "resetPassword") {
    var identificativo = e.parameter.telefono.trim();
    var dataDB = sheetDB.getDataRange().getValues();
    for (var r = 1; r < dataDB.length; r++) {
      if (dataDB[r][0].toString().trim().toLowerCase() === identificativo.toLowerCase()) {

        // 1. Prende il cellulare dalla Colonna E (indice 4)
        var passwordIniziale = dataDB[r][4];

        // 2. Lo copia nella Colonna C (indice 3)
        sheetDB.getRange(r + 1, 3).setValue(passwordIniziale);

        // 3. SVUOTA la Colonna D (indice 4) per far scattare il "deveCambiare" al prossimo login
        sheetDB.getRange(r + 1, 4).setValue("");

        return successo({ resetInviato: true });
      }
    }
    return successo({ resetInviato: false });
  }
}

  function successo(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  }