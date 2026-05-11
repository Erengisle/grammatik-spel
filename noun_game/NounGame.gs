/**
 * NounGame.gs — Substantivspelet (Dokument 2)
 *
 * Separat Google Apps Script-projekt för gamifierad substantivövning.
 * Elever övar genus (en/ett) och får poäng för rätta svar.
 *
 * Blad i detta dokument:
 *   Ordlista  – ord, genus, plural, deklination, saldo_paradigm, aktiv
 *   Poäng     – elev, timestamp, poäng, rätt, totalt
 *
 * Kom igång:
 *   1. Skapa ett nytt Google Sheets-dokument
 *   2. Tillägg → Apps Script → klistra in NounGame.gs
 *   3. Lägg till NounGame.html som en HTML-fil
 *   4. Kör: Spelverktyg → Initiera Ordlista-blad
 *   5. Kör: Spelverktyg → Importera från Kelly-listan
 *   6. Kör: Spelverktyg → Berika med genus via SALDO  (för ord som saknar genus)
 *   7. Publicera → Distribuera som webbapp
 *
 * Kelly-lista (källa): CC-BY-SA, Göteborgs universitet / Språkbanken
 */

// ID och fliks-GID för Kelly-listan (ditt Google Sheets-dokument)
var KELLY_SHEET_ID  = '1G2B06J0cHSdhj5BMxBvdZui2YI7UFxDglBBoGmfTHxM';
var KELLY_SHEET_GID = 302703246;

// ─── Webb-ingång ───────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('NounGame')
    .setTitle('Substantivspelet')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Klient-API ────────────────────────────────────────────────────────────────

/**
 * Returnerar alla aktiva spelord till klienten.
 * Filtrerar bort ord utan genus.
 */
function hamtaSpelOrd() {
  var ss     = SpreadsheetApp.getActive();
  var sheet  = ss.getSheetByName('Ordlista');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return h.toString().toLowerCase(); });
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var ordIdx   = headers.indexOf('ord');
  var genusIdx = headers.indexOf('genus');
  var aktivIdx = headers.indexOf('aktiv');

  if (ordIdx < 0 || genusIdx < 0) return [];

  return data
    .filter(function(r) {
      var aktiv = aktivIdx >= 0 ? (r[aktivIdx] || '').toString().toLowerCase() : 'ja';
      return aktiv !== 'nej';
    })
    .map(function(r) {
      return {
        ord:   (r[ordIdx]   || '').toString().trim(),
        genus: (r[genusIdx] || '').toString().toLowerCase().trim()
      };
    })
    .filter(function(r) {
      return r.ord && (r.genus === 'en' || r.genus === 'ett');
    });
}

/**
 * Sparar en elevs rundresultat i Poäng-bladet.
 */
function sparaPoang(payload) {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Poäng');

  if (!sheet) {
    sheet = ss.insertSheet('Poäng');
    sheet.appendRow(['Elev', 'Timestamp', 'Poäng', 'Rätt', 'Totalt']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    (payload.student || '').toString(),
    new Date(),
    Number(payload.score)   || 0,
    Number(payload.correct) || 0,
    Number(payload.total)   || 0
  ]);
}

// ─── Meny ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Spelverktyg')
    .addItem('Initiera Ordlista-blad',              'initieraOrdlista')
    .addItem('Importera från Kelly-listan',         'importeraFranKelly')
    .addItem('Importera från Ordbank (Dokument 1)', 'importeraFranOrdbank')
    .addSeparator()
    .addItem('Berika med genus via SALDO',          'berikaMedSaldoGenus')
    .addItem('Testa SALDO för ett ord',             'testSaldoOrd')
    .addToUi();
}

// ─── Ordlista ──────────────────────────────────────────────────────────────────

/**
 * Skapar Ordlista-bladet med rätt rubriker.
 * Befintliga data skrivs INTE över om bladet redan finns.
 */
function initieraOrdlista() {
  var ss       = SpreadsheetApp.getActive();
  var sheet    = ss.getSheetByName('Ordlista');
  var nyskapad = false;

  if (!sheet) { sheet = ss.insertSheet('Ordlista'); nyskapad = true; }

  if (nyskapad || sheet.getLastRow() === 0) {
    var headers = ['ord', 'genus', 'plural', 'deklination', 'saldo_paradigm', 'aktiv'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  SpreadsheetApp.getActive().toast('Ordlista-bladet är klart.', 'Klart', 4);
}

// ─── Kelly-import ──────────────────────────────────────────────────────────────

/**
 * Importerar substantiv från Kelly-listan (Google Sheets-dokument).
 * Kolumn A = genus (en/ett eller tomt), Kolumn B = grundform.
 */
function importeraFranKelly() {
  var kellySS;
  try {
    kellySS = SpreadsheetApp.openById(KELLY_SHEET_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Kunde inte öppna Kelly-dokumentet.\n\n' +
      'Kontrollera att dokumentet är delat med samma Google-konto.\n\nFel: ' + e.message
    );
    return;
  }

  var kellySh   = null;
  var allSheets = kellySS.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getSheetId() === KELLY_SHEET_GID) { kellySh = allSheets[i]; break; }
  }
  if (!kellySh) kellySh = allSheets[0];

  var lastRow = kellySh.getLastRow();
  if (lastRow < 1) { SpreadsheetApp.getUi().alert('Kelly-fliken verkar vara tom.'); return; }

  var raw = kellySh.getRange(1, 1, lastRow, 2).getValues();

  initieraOrdlista();
  var dest    = SpreadsheetApp.getActive().getSheetByName('Ordlista');
  var headers = dest.getRange(1, 1, 1, dest.getLastColumn()).getValues()[0]
    .map(function(h) { return h.toString().toLowerCase(); });

  var iOrd   = headers.indexOf('ord');
  var iGenus = headers.indexOf('genus');
  var iAktiv = headers.indexOf('aktiv');

  var befintliga = {};
  var destLast   = dest.getLastRow();
  if (destLast >= 2) {
    dest.getRange(2, iOrd + 1, destLast - 1, 1).getValues()
      .forEach(function(r) { befintliga[(r[0] || '').toString().toLowerCase()] = true; });
  }

  var nyaRader = [], medGenus = 0, utanGenus = 0, dubbletter = 0;

  raw.forEach(function(row) {
    var genus = (row[0] || '').toString().trim().toLowerCase();
    var ord   = (row[1] || '').toString().trim().toLowerCase();
    if (!ord) return;
    if (befintliga[ord]) { dubbletter++; return; }
    befintliga[ord] = true;

    var genusVal = (genus === 'en' || genus === 'ett') ? genus : '';
    var rad      = new Array(headers.length).fill('');
    if (iOrd   >= 0) rad[iOrd]   = ord;
    if (iGenus >= 0) rad[iGenus] = genusVal;
    if (iAktiv >= 0) rad[iAktiv] = 'ja';
    nyaRader.push(rad);
    if (genusVal) medGenus++; else utanGenus++;
  });

  if (nyaRader.length === 0) {
    SpreadsheetApp.getActive().toast('Inga nya ord att importera.', 'Kelly-import', 5);
    return;
  }

  dest.getRange(dest.getLastRow() + 1, 1, nyaRader.length, headers.length).setValues(nyaRader);

  SpreadsheetApp.getActive().toast(
    nyaRader.length + ' ord importerade  (' + dubbletter + ' dubbletter hoppades över)\n' +
    '✓ Med genus: ' + medGenus + '   ? Saknar genus: ' + utanGenus + '\n' +
    'Kör "Berika med genus via SALDO" för att fylla i de som saknas.',
    'Kelly-import klar', 12
  );
}

// ─── Ordbank-import ────────────────────────────────────────────────────────────

/**
 * Importerar substantiv från Ordbanken i Dokument 1.
 */
function importeraFranOrdbank() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Importera substantiv',
    'Klistra in Google Sheets-ID för dokumentet med Ordbanken.\n' +
    'ID:t finns i URL:en: docs.google.com/spreadsheets/d/[ID]/edit',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() === ui.Button.CANCEL) return;

  var sourceId = resp.getResponseText().trim();
  if (!sourceId) { ui.alert('Inget ID angavs.'); return; }

  try {
    var sourceSS    = SpreadsheetApp.openById(sourceId);
    var sourceSheet = sourceSS.getSheetByName('Ordbank') || sourceSS.getSheetByName('Översikt');
    if (!sourceSheet) { ui.alert('Hittade inte bladet "Ordbank" eller "Översikt".'); return; }

    var values = sourceSheet.getDataRange().getValues();
    if (values.length < 2) { ui.alert('Källbladet saknar data.'); return; }

    var srcH     = values[0].map(function(h) { return h.toString().toLowerCase(); });
    var lemmaIdx = srcH.indexOf('lemma');
    var okIdx    = srcH.indexOf('ordklass');
    var deklIdx  = srcH.indexOf('deklination');
    if (lemmaIdx < 0) { ui.alert('Källbladet saknar kolumnen "lemma".'); return; }

    var substantiv = values.slice(1).filter(function(r) {
      var ok = okIdx >= 0 ? (r[okIdx] || '').toString().toLowerCase() : '';
      return ok === 'substantiv' || ok === '';
    });

    initieraOrdlista();
    var dest    = SpreadsheetApp.getActive().getSheetByName('Ordlista');
    var destH   = dest.getRange(1, 1, 1, dest.getLastColumn()).getValues()[0]
      .map(function(h) { return h.toString().toLowerCase(); });

    var dOrd   = destH.indexOf('ord');
    var dDekl  = destH.indexOf('deklination');
    var dAktiv = destH.indexOf('aktiv');

    var befintliga = {};
    if (dest.getLastRow() >= 2) {
      dest.getRange(2, dOrd + 1, dest.getLastRow() - 1, 1).getValues()
        .forEach(function(r) { befintliga[(r[0] || '').toString().toLowerCase()] = true; });
    }

    var nyaRader = substantiv
      .filter(function(r) { return !befintliga[(r[lemmaIdx] || '').toString().toLowerCase()]; })
      .map(function(r) {
        var rad = new Array(destH.length).fill('');
        if (dOrd   >= 0) rad[dOrd]   = (r[lemmaIdx] || '').toString();
        if (dDekl  >= 0) rad[dDekl]  = deklIdx >= 0 ? (r[deklIdx] || '').toString() : '';
        if (dAktiv >= 0) rad[dAktiv] = 'ja';
        return rad;
      });

    if (nyaRader.length > 0) {
      dest.getRange(dest.getLastRow() + 1, 1, nyaRader.length, destH.length).setValues(nyaRader);
    }

    SpreadsheetApp.getActive().toast(
      nyaRader.length + ' nya substantiv importerade.\nKör nu "Berika med genus via SALDO".',
      'Import klar', 8
    );
  } catch (err) {
    ui.alert('Fel vid import: ' + err.message);
  }
}

// ─── SALDO-berikning ───────────────────────────────────────────────────────────

var SALDO_WS = 'https://spraakbanken.gu.se/ws/saldo-ws';

/**
 * Visar råsvaret från SALDO för ett enskilt ord — för felsökning.
 */
function testSaldoOrd() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt('Testa SALDO', 'Vilket ord vill du testa?', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.CANCEL) return;

  var word = resp.getResponseText().trim().toLowerCase();
  var url  = SALDO_WS + '/fl/json/' + encodeURIComponent(word);

  try {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    ui.alert(
      'SALDO /fl/json/' + word +
      '\nHTTP: ' + r.getResponseCode() + '\n\n' +
      r.getContentText('UTF-8').substring(0, 800)
    );
  } catch (e) {
    ui.alert('Fel: ' + e.message);
  }
}

/**
 * Slår upp ett ord i SALDO via /fl/json/{ord}.
 * Väljer posten där gf exakt matchar sökordet (filtrerar bort sammansättningar).
 * Paradigmet finns i fältet "p".
 */
function saldoSlaSuppOrd_(word) {
  var url = SALDO_WS + '/fl/json/' + encodeURIComponent(word);
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { Accept: 'application/json' }
    });
    if (resp.getResponseCode() !== 200) return { hittad: false };

    var data = JSON.parse(resp.getContentText('UTF-8'));
    if (!Array.isArray(data) || data.length === 0) return { hittad: false };

    // Hitta posten där grundformen (gf) exakt matchar sökordet
    var entry = null;
    for (var i = 0; i < data.length; i++) {
      if ((data[i].gf || '').toLowerCase().trim() === word) { entry = data[i]; break; }
    }
    if (!entry) return { hittad: false };

    var paradigm = entry.p || '';
    var prefix   = paradigm.split('_')[0];
    var ordklass = { vb: 'verb', nn: 'substantiv', jj: 'adjektiv' }[prefix] || prefix;

    return { hittad: true, paradigm: paradigm, ordklass: ordklass };
  } catch (e) {
    return { hittad: false };
  }
}

// nn_Xu_* → en   nn_Xn_* → ett
function tolkaSaldoGenus_(paradigm) {
  if (!paradigm || paradigm.indexOf('nn_') !== 0) return '';
  var grp = paradigm.split('_')[1] || '';
  if (grp.charAt(grp.length - 1) === 'u') return 'en';
  if (grp.charAt(grp.length - 1) === 'n') return 'ett';
  return '';
}

// nn_1u_* → 1 | nn_2u_* → 2 | nn_3u_* → 3 | nn_4n_* → 4 | nn_5n_* → 5
function tolkaSaldoDeklination_(paradigm) {
  if (!paradigm || paradigm.indexOf('nn_') !== 0) return '';
  return (paradigm.split('_')[1] || '').replace(/[un]$/, '');
}

/**
 * Fyller i genus, deklination och saldo_paradigm för ord utan genus
 * i Ordlistan via SALDO-uppslag.
 */
function berikaMedSaldoGenus() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Ordlista');
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Ordlistan är tom.');
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return h.toString().toLowerCase(); });

  var ordIdx      = headers.indexOf('ord');
  var genusIdx    = headers.indexOf('genus');
  var paradigmIdx = headers.indexOf('saldo_paradigm');
  var deklIdx     = headers.indexOf('deklination');

  if (ordIdx < 0 || genusIdx < 0) {
    SpreadsheetApp.getUi().alert('Ordlistan saknar kolumnerna "ord" eller "genus".');
    return;
  }

  var data        = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var uppdaterade = 0, ejHittade = 0;

  for (var i = 0; i < data.length; i++) {
    var ord      = (data[i][ordIdx]   || '').toString().trim().toLowerCase();
    var harGenus = (data[i][genusIdx] || '').toString().trim();
    if (!ord || harGenus === 'en' || harGenus === 'ett') continue;

    if (i % 15 === 0 && i > 0) {
      SpreadsheetApp.getActive().toast('Bearbetar rad ' + (i + 2) + '…', 'SALDO', 2);
    }
    Utilities.sleep(200);

    var res = saldoSlaSuppOrd_(ord);

    if (res.hittad && res.ordklass === 'substantiv') {
      var genus = tolkaSaldoGenus_(res.paradigm);
      var dekl  = tolkaSaldoDeklination_(res.paradigm);
      if (genus) {
        sheet.getRange(i + 2, genusIdx + 1).setValue(genus);
        if (paradigmIdx >= 0) sheet.getRange(i + 2, paradigmIdx + 1).setValue(res.paradigm);
        if (deklIdx    >= 0) sheet.getRange(i + 2, deklIdx    + 1).setValue(dekl);
        uppdaterade++;
      } else {
        ejHittade++;
      }
    } else {
      ejHittade++;
    }
  }

  SpreadsheetApp.getActive().toast(
    'Genus ifyllt: ' + uppdaterade + '   Ej hittade / okänt genus: ' + ejHittade,
    'SALDO-berikning klar', 8
  );
}
