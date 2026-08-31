/**
 * Geomad Marketing Hub — ponte con Google Sheets
 *
 * Riceve i dati inviati dal prototipo (una o più "schede" con intestazioni e righe)
 * e le scrive nel foglio Google, sostituendo ogni volta il contenuto della scheda con
 * i dati più recenti (nessun accumulo di righe duplicate ad ogni nuovo caricamento).
 * Espone anche una lettura (doGet) di tutte le schede, così chiunque apra la pagina
 * del prototipo vede la stessa lavorazione, senza dover ricaricare i file.
 *
 * Gestisce anche l'accesso individuale (login): le credenziali si trovano nella
 * scheda "Utenti" del foglio, che questo script NON restituisce mai tramite doGet —
 * quindi le password non compaiono mai tra i dati che il prototipo legge normalmente.
 *
 * INSTALLAZIONE
 * 1. Apri (o crea) il Google Sheet che vuoi usare come destinazione.
 * 2. Menu Estensioni > Apps Script (se questa voce non è disponibile, vedi la nota
 *    più sotto: puoi comunque creare lo script da script.google.com).
 * 3. Cancella il contenuto di default e incolla tutto questo file.
 * 4. Se hai creato lo script da Estensioni > Apps Script (agganciato al foglio),
 *    lascia SHEET_ID vuoto: lo script trova da solo il foglio a cui è agganciato.
 *    Se invece hai creato lo script da script.google.com (script "standalone", non
 *    agganciato a nessun foglio), incolla qui sotto l'ID del foglio: lo trovi
 *    nell'URL del foglio, tra /d/ e /edit — es. da
 *    https://docs.google.com/spreadsheets/d/ABC123.../edit copi "ABC123...".
 * 5. Salva (icona del dischetto, o Ctrl/Cmd+S). Dai un nome al progetto se richiesto.
 * 6. In alto a destra: Distribuisci > Nuova distribuzione.
 * 7. Icona a ingranaggio accanto a "Seleziona tipo" > "App web".
 * 8. Configurazione: "Esegui come" = Me (il tuo account); "Chi ha accesso" = Chiunque.
 * 9. Distribuisci. Google chiederà di autorizzare lo script al primo utilizzo: accetta
 *    (comparirà un avviso "app non verificata" perché è tuo/personale — è normale,
 *    scegli "Avanzate" > "Vai al progetto (non sicuro)" per procedere).
 * 10. Copia l'URL dell'app web che ti viene mostrato: è quello da incollare nel campo
 *     "Sincronizzazione Google Sheets" del prototipo.
 * 11. Se in futuro modifichi questo file (es. per aggiornarlo), usa "Gestisci
 *     distribuzioni > Modifica > Nuova versione" per mantenere lo stesso URL —
 *     altrimenti dovresti incollare un URL nuovo nel prototipo.
 *
 * PER ATTIVARE IL LOGIN
 * 12. Nel foglio Google, crea una nuova scheda chiamata esattamente "Utenti".
 * 13. Nella riga 1 metti le intestazioni: Nome | Email | Password | Ruolo | Accesso
 * 14. Da riga 2 in poi, una riga per persona, es.:
 *     Mara Figus | mara.figus@geomad.site | unapasswordascelta | Admin | Caricamento
 *     L'email è quella che la persona userà per accedere alla Marketing Hub, non
 *     deve essere per forza un indirizzo vero.
 * 15. Non serve nessuna distribuzione aggiuntiva: la scheda "Utenti" viene letta
 *     automaticamente da questo stesso script quando qualcuno prova ad accedere.
 *
 * RUOLO decide quali pagine vede la persona una volta entrata. Puoi scrivere più
 * ruoli nella stessa cella separandoli con una virgola (es. "Meta, Social": vede
 * entrambe le sezioni). I valori possibili sono:
 *   - Admin: vede e può caricare tutto, incluso "Utenti & permessi". Se "Admin"
 *     compare tra i valori (anche insieme ad altri), vince sempre lui.
 *   - Google: vede la sezione Google.
 *   - Meta: vede la sezione Meta (costi teleselling, dettagli, campagne...).
 *   - Social: vede la sezione Social.
 *   - Grafiche: vede la sezione Grafiche.
 *   Un Ruolo vuoto, o dove nessun valore è scritto correttamente, viene trattato
 *   come Admin — così un refuso non blocca nessuno fuori per errore.
 *
 * ACCESSO decide se la persona può anche caricare/modificare o solo guardare:
 *   - Caricamento (o vuoto): può caricare i file e modificare i dati della sua sezione.
 *   - Lettura: vede solo i dati, i pulsanti e i campi per caricare o modificare sono
 *     nascosti. Non si applica ad Admin, che ha sempre accesso completo.
 *
 * NOTE DI SICUREZZA: con "Chiunque" come accesso, chiunque conosca questo URL può
 * leggere e scrivere i dati nelle schede che il prototipo usa (tranne "Utenti", mai
 * restituita). Il login impedisce l'accesso occasionale tramite la pagina del sito,
 * ma non sostituisce una vera sicurezza informatica: non usarlo per dati davvero
 * riservati, e non riusare per la scheda "Utenti" password importanti usate altrove.
 * Non condividere questo URL pubblicamente: trattalo come condivideresti il foglio.
 */

// Lascia vuoto se lo script è agganciato a un foglio (creato da Estensioni > Apps
// Script). Incolla l'ID del foglio solo se lo script è "standalone" (creato da
// script.google.com) — vedi il punto 4 delle istruzioni qui sopra.
var SHEET_ID = '';

// Nome della scheda con le credenziali di accesso — vedi i punti 12-15 qui sopra.
var USERS_SHEET_NAME = 'Utenti';

function getSpreadsheet() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = getSpreadsheet();
    if (payload.action === 'login') {
      return handleLogin(ss, payload.username, payload.password);
    }
    var sheets = payload.sheets || [];
    sheets.forEach(function (s) {
      if (s) writeSheet(ss, s.name, s.headers, s.rows, s.textCols);
    });
    return jsonOutput({ ok: true, updated: sheets.length });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

// Controlla email + password nella scheda "Utenti" (Nome, Email, Password, Ruolo).
// Non rivela mai se è l'email o la password ad essere sbagliata, e non restituisce
// mai l'elenco delle password: solo ok true/false e, se ok, il nome (e il ruolo, se
// presente) da mostrare nella pagina.
function handleLogin(ss, username, password) {
  var sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    return jsonOutput({ ok: false, error: 'missing_users_sheet' });
  }
  var values = sheet.getDataRange().getValues();
  var uname = String(username || '').trim().toLowerCase();
  var pass = String(password || '');
  if (!uname || !pass) return jsonOutput({ ok: false });
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowEmail = String(row[1] || '').trim().toLowerCase();
    var rowPass = String(row[2] || '');
    if (rowEmail && rowEmail === uname && rowPass === pass) {
      return jsonOutput({ ok: true, name: row[0] || username, role: row[3] || '', access: row[4] || '' });
    }
  }
  return jsonOutput({ ok: false });
}

// Restituisce il contenuto di tutte le schede del foglio (tranne "Utenti", che non
// viene mai esposta), così il prototipo può ricostruire la lavorazione più recente
// all'apertura della pagina o quando si preme "Aggiorna dal foglio Google" — senza
// che nessuno debba ricaricare i file.
function doGet(e) {
  try {
    var ss = getSpreadsheet();
    var sheets = ss.getSheets();
    var out = {};
    sheets.forEach(function (sheet) {
      if (sheet.getName() === USERS_SHEET_NAME) return;
      var values = sheet.getDataRange().getValues();
      if (!values.length) {
        out[sheet.getName()] = { headers: [], rows: [] };
        return;
      }
      out[sheet.getName()] = { headers: values[0], rows: values.slice(1) };
    });
    return jsonOutput({ ok: true, sheets: out });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

// textCols: indici di colonna (0 = prima colonna) da scrivere come testo puro, così
// Google Sheets non li reinterpreta a modo suo — capita soprattutto con le date
// (es. "2026-08-27" convertito in una data vera e poi riletto con un fuso orario
// diverso può spostarsi di un giorno) e con gli ID campagna molto lunghi (numeri a
// 18 cifre che, se trattati come numeri, perdono le ultime cifre per arrotondamento).
function writeSheet(ss, name, headers, rows, textCols) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  sheet.clearContents();
  sheet.clearFormats();
  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (rows && rows.length) {
    if (textCols && textCols.length) {
      textCols.forEach(function (colIdx0) {
        if (colIdx0 >= 0 && colIdx0 < headers.length) {
          sheet.getRange(2, colIdx0 + 1, rows.length, 1).setNumberFormat('@');
        }
      });
    }
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

/** Utile per un test rapido da dentro l'editor Apps Script (Esegui > testDoPost). */
function testDoPost() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        sheets: [
          {
            name: 'Test',
            headers: ['Colonna A', 'Colonna B'],
            rows: [
              ['ciao', 1],
              ['mondo', 2]
            ]
          }
        ]
      })
    }
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}

/** Utile per un test rapido da dentro l'editor Apps Script (Esegui > testDoGet). */
function testDoGet() {
  var result = doGet({});
  Logger.log(result.getContent());
}

/**
 * Utile per un test rapido del login da dentro l'editor Apps Script
 * (Esegui > testLogin). Modifica email/password con una riga vera della scheda
 * "Utenti" prima di eseguirlo.
 */
function testLogin() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        action: 'login',
        username: 'mara.figus@geomad.site',
        password: 'scrivi-qui-la-password-di-prova'
      })
    }
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
