// @ts-nocheck
//Versione 
// @ts-nocheck Versione 3.4
function importaEventiCalendarioAvanzato() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var foglioDestinazione = ss.getSheetByName("TAM");
  var sheetDB = ss.getSheetByName("Database_PW");

  // --- 1. PULIZIA INIZIALE DEL FOGLIO TAM (SOLO UNA VOLTA) ---
  if (foglioDestinazione.getLastRow() > 1) {
    foglioDestinazione.getRange(2, 1, foglioDestinazione.getLastRow(), 3).clearContent();
  }

  // --- 2. RECUPERO DATI DAL DATABASE_PW ---
  var datiDB = sheetDB.getDataRange().getValues();
  
  // --- 3. IMPOSTAZIONI TEMPORALI ---
  var inizio = new Date();
  inizio.setHours(0, 0, 0, 0);
  var fine = new Date();
  fine.setMonth(inizio.getMonth() + 2);

  // --- 4. CICLO SULLE RIGHE DEL DATABASE ---
  for (var r = 1; r < datiDB.length; r++) {
    var nomeDaScrivere = datiDB[r][0]; // Colonna A (Indice 0)
    var calId = datiDB[r][9];         // Colonna J (Indice 9)

    if (calId && calId.toString().trim() !== "") {
      try {
        var calendario = CalendarApp.getCalendarById(calId.toString().trim());
        if (calendario) {
          var eventi = calendario.getEvents(inizio, fine);
          var outputTemporaneo = [];

          for (var j = 0; j < eventi.length; j++) {
            var ev = eventi[j];
            var oraInizio = Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "HH:mm");
            var oraFine = Utilities.formatDate(ev.getEndTime(), Session.getScriptTimeZone(), "HH:mm");
            var descrizioneCompleta = "TAM - " + ev.getTitle() + " " + oraInizio + " - " + oraFine;

            outputTemporaneo.push([
              ev.getStartTime(),
              descrizioneCompleta,
              nomeDaScrivere // <--- Ora usa il nome preso dalla Colonna A del Database_PW
            ]);
          }

          // Scrittura in coda per ogni calendario processato
          if (outputTemporaneo.length > 0) {
            var rigaPartenza = foglioDestinazione.getLastRow() + 1;
            foglioDestinazione.getRange(rigaPartenza, 1, outputTemporaneo.length, 3).setValues(outputTemporaneo);
            // Formattazione data
            foglioDestinazione.getRange(rigaPartenza, 1, outputTemporaneo.length, 1).setNumberFormat("dd/mm/yyyy");
          }
        }
      } catch (e) {
        Logger.log("Errore con il calendario di " + nomeDaScrivere + " (" + calId + "): " + e.message);
      }
    }
  }
}

function doGet(e) {
  var azione = e.parameter.azione;
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Database_PW"); // Assicurati che il nome del foglio sia corretto
  
  // GESTIONE DELLA RICERCA UTENTE
  if (azione === "cerca") {
    var ricerca = e.parameter.email.toLowerCase().trim();
    var dati = sheet.getDataRange().getValues(); // Legge tutta la tabella
    
    for (var i = 1; i < dati.length; i++) {
      var nominativo = String(dati[i][0]).toLowerCase().trim(); // Colonna A (Nominativo)
      var email = String(dati[i][1]).toLowerCase().trim();     // Colonna B (Telefono/Email)
      var passwordColC = String(dati[i][2]).trim();             // Colonna C (Password / Stato)
      
      // Controlla se corrisponde al nome o all'email inseriti
      if (nominativo === ricerca || email === ricerca) {
        
        // CONTROLLO BLOCCO: Se la colonna C contiene "NoAcces"
        if (passwordColC === "NoAcces") {
          return ContentService.createTextOutput(JSON.stringify({
            trovato: false,
            errore: "UTENTE NON ABILITATO ALL'ACCESSO"
          })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // Se l'utente è abilitato, restituisce i dati normali
        return ContentService.createTextOutput(JSON.stringify({
          trovato: true,
          nome: dati[i][0],
          email: dati[i][1]
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // Se non trova nessuno
    return ContentService.createTextOutput(JSON.stringify({
      trovato: false
    })).setMimeType(ContentService.MimeType.JSON);
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

  // --- NUOVA AZIONE: RESET PASSWORD ---
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
    } // Chiude il ciclo for
    return successo({ resetInviato: false }); // Ritorna falso se l'utente non viene trovato
  } // Chiude il blocco if (azione === "resetPassword")
} // <--- QUESTA CHIUDE LA FUNZIONE doGet(e)

function successo(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}