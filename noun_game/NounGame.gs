/**
 * NounGame.gs — Substantivspelet (Dokument 2)
 *
 * Blad i detta dokument:
 *   Ordlista  – ord, genus, plural, deklination, saldo_paradigm, cefr, aktiv
 *   Poäng     – elev, timestamp, poäng, rätt, totalt
 *
 * Kom igång:
 *   1. Skapa ett nytt Google Sheets-dokument
 *   2. Tillägg → Apps Script → klistra in NounGame.gs
 *   3. Lägg till NounGame.html och Highscore.html som HTML-filer
 *   4. Spelverktyg → Initiera Ordlista-blad
 *   5. Spelverktyg → Importera från Kelly-listan
 *   6. Spelverktyg → Berika med genus via SALDO
 *   7. Spelverktyg → Berika med plural via SALDO
 *   8. Publicera → Distribuera som webbapp
 *
 * Kelly-lista (källa): CC-BY-SA, Göteborgs universitet / Språkbanken
 */

var KELLY_SHEET_ID   = '1G2B06J0cHSdhj5BMxBvdZui2YI7UFxDglBBoGmfTHxM';
var KELLY_SHEET_GID  = 302703246;
var ORDLISTA_HEADERS = ['ord', 'genus', 'plural', 'deklination', 'saldo_paradigm', 'cefr', 'aktiv'];

// ─── Webb-ingång ─────────────────────────────────────────────────────────────

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'highscore') {
    return HtmlService.createHtmlOutputFromFile('Highscore')
      .setTitle('Highscore — Substantivspelet')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('NounGame')
    .setTitle('Substantivspelet')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Klient-API ──────────────────────────────────────────────────────────────

/**
 * Returnerar aktiva spelord för genus-övning.
 * @param {{ cefr?: string }} options
 * @returns {Array<{ord, genus, cefr}>}
 */
function hamtaSpelOrd(options) {
  var cefrFilter = (options && options.cefr) ? options.cefr.toUpperCase() : '';
  var sheet = _getOrdlista();
  if (!sheet) return [];

  var idx = _getOrdlistaIndex(sheet);
  if (idx.ord < 0 || idx.genus < 0) return [];

  return _getOrdlistaData(sheet)
    .filter(function(r) {
      var aktiv = idx.aktiv >= 0 ? (r[idx.aktiv] || '').toString().toLowerCase() : 'ja';
      if (aktiv === 'nej') return false;
      if (cefrFilter && idx.cefr >= 0) {
        var wCefr = (r[idx.cefr] || '').toString().toUpperCase().trim();
        if (wCefr && wCefr !== cefrFilter) return false;
      }
      return true;
    })
    .map(function(r) {
      return {
        ord:   (r[idx.ord]   || '').toString().trim(),
        genus: (r[idx.genus] || '').toString().toLowerCase().trim(),
        cefr:  idx.cefr >= 0 ? (r[idx.cefr] || '').toString().toUpperCase().trim() : ''
      };
    })
    .filter(function(r) {
      return r.ord && (r.genus === 'en' || r.genus === 'ett');
    });
}

/**
 * Returnerar aktiva spelord för plural-övning.
 * @param {{ cefr?: string }} options
 * @returns {Array<{ord, genus, plural, cefr}>}
 */
function hamtaSpelOrdPlural(options) {
  var cefrFilter = (options && options.cefr) ? options.cefr.toUpperCase() : '';
  var sheet = _getOrdlista();
  if (!sheet) return [];

  var idx = _getOrdlistaIndex(sheet);
  if (idx.ord < 0 || idx.plural < 0) return [];

  return _getOrdlistaData(sheet)
    .filter(function(r) {
      var aktiv = idx.aktiv >= 0 ? (r[idx.aktiv] || '').toString().toLowerCase() : 'ja';
      if (aktiv === 'nej') return false;
      if (cefrFilter && idx.cefr >= 0) {
        var wCefr = (r[idx.cefr] || '').toString().toUpperCase().trim();
        if (wCefr && wCefr !== cefrFilter) return false;
      }
      return (r[idx.plural] || '').toString().trim() !== '';
    })
    .map(function(r) {
      return {
        ord:    (r[idx.ord]    || '').toString().trim(),
        genus:  idx.genus >= 0 ? (r[idx.genus] || '').toString().toLowerCase().trim() : '',
        plural: (r[idx.plural] || '').toString().trim().toLowerCase(),
        cefr:   idx.cefr >= 0 ? (r[idx.cefr] || '').toString().toUpperCase().trim() : ''
      };
    })
    .filter(function(r) { return r.ord && r.plural; });
}

/**
 * Returnerar highscore-data aggregerat per elev + 20 senaste omgångar.
 * @returns {{ leaderboard: Array, recentGames: Array }}
 */
function hamtaHighscore() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Poäng');
  if (!sheet || sheet.getLastRow() < 2) return { leaderboard: [], recentGames: [] };

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  var byStudent = {};
  data.forEach(function(r) {
    var student = (r[0] || '').toString().trim();
    if (!student) return;
    if (!byStudent[student]) {
      byStudent[student] = { student: student, totalScore: 0, totalCorrect: 0, totalTotal: 0, games: 0 };
    }
    byStudent[student].totalScore   += Number(r[2]) || 0;
    byStudent[student].totalCorrect += Number(r[3]) || 0;
    byStudent[student].totalTotal   += Number(r[4]) || 0;
    byStudent[student].games++;
  });

  var leaderboard = Object.keys(byStudent)
    .map(function(k) { return byStudent[k]; })
    .sort(function(a, b) { return b.totalScore - a.totalScore; })
    .slice(0, 25);

  var recentGames = data.slice(-20).reverse().map(function(r) {
    return {
      student:   (r[0] || '').toString().trim(),
      timestamp: r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '',
      score:     Number(r[2]) || 0,
      correct:   Number(r[3]) || 0,
      total:     Number(r[4]) || 0
    };
  });

  return { leaderboard: leaderboard, recentGames: recentGames };
}

/**
 * Sparar en elevs rundresultat i Poäng-bladet.
 * @param {{ student, score, correct, total }} payload
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

// ─── Admin ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Spelverktyg')
    .addItem('Initiera Ordlista-blad',              'initieraOrdlista')
    .addItem('Importera från Kelly-listan',         'importeraFranKelly')
    .addItem('Importera från Ordbank (Dokument 1)', 'importeraFranOrdbank')
    .addSeparator()
    .addItem('Berika med genus via SALDO',          'berikaMedSaldoGenus')
    .addItem('Berika med plural via SALDO',         'berikaMedSaldoPlural')
    .addToUi();
}

// ─── Interna hjälpfunktioner ──────────────────────────────────────────────────

function _getOrdlista() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ordlista');
  return (sheet && sheet.getLastRow() >= 2) ? sheet : null;
}

function _getOrdlistaIndex(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return h.toString().toLowerCase(); });
  return {
    ord:         headers.indexOf('ord'),
    genus:       headers.indexOf('genus'),
    plural:      headers.indexOf('plural'),
    deklination: headers.indexOf('deklination'),
    paradigm:    headers.indexOf('saldo_paradigm'),
    cefr:        headers.indexOf('cefr'),
    aktiv:       headers.indexOf('aktiv')
  };
}

function _getOrdlistaData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

// ─── Ordlista-initiering ──────────────────────────────────────────────────────

/**
 * Skapar Ordlista-bladet med rätt rubriker.
 * Lägger till saknade kolumner (t.ex. cefr) utan att röra befintlig data.
 */
function initieraOrdlista() {
  var ss       = SpreadsheetApp.getActive();
  var sheet    = ss.getSheetByName('Ordlista');
  var nyskapad = false;

  if (!sheet) {
    sheet    = ss.insertSheet('Ordlista');
    nyskapad = true;
  }

  if (nyskapad || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ORDLISTA_HEADERS.length)
      .setValues([ORDLISTA_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, ORDLISTA_HEADERS.length);
  } else {
    // Lägg till saknade kolumner utan att röra befintlig data
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return h.toString().toLowerCase(); });
    ORDLISTA_HEADERS.forEach(function(h) {
      if (existing.indexOf(h.toLowerCase()) < 0) {
        var col = sheet.getLastColumn() + 1;
        sheet.getRange(1, col).setValue(h).setFontWeight('bold');
        sheet.autoResizeColumn(col);
        existing.push(h.toLowerCase());
      }
    });
  }

  ss.toast('Ordlista-bladet är klart.', 'Klart', 6);
}

// ─── Kelly-import ─────────────────────────────────────────────────────────────

/**
 * Importerar substantiv från Kelly-listan.
 * Kolumn A = genus, Kolumn B = grundform. Kolumn D = CEFR om den finns.
 */
function importeraFranKelly() {
  var kellySS;
  try {
    kellySS = SpreadsheetApp.openById(KELLY_SHEET_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Kunde inte öppna Kelly-dokumentet.\n\nFel: ' + e.message
    );
    return;
  }

  var kellySh  = null;
  var allSheets = kellySS.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getSheetId() === KELLY_SHEET_GID) { kellySh = allSheets[i]; break; }
  }
  if (!kellySh) kellySh = allSheets[0];

  var lastRow = kellySh.getLastRow();
  if (lastRow < 1) { SpreadsheetApp.getUi().alert('Kelly-fliken verkar vara tom.'); return; }

  var numCols = Math.min(kellySh.getLastColumn(), 4);
  var raw     = kellySh.getRange(1, 1, lastRow, numCols).getValues();

  initieraOrdlista();
  var dest    = SpreadsheetApp.getActive().getSheetByName('Ordlista');
  var idx     = _getOrdlistaIndex(dest);
  var headers = dest.getRange(1, 1, 1, dest.getLastColumn()).getValues()[0]
    .map(function(h) { return h.toString().toLowerCase(); });

  var befintliga = {};
  var destLast   = dest.getLastRow();
  if (destLast >= 2 && idx.ord >= 0) {
    dest.getRange(2, idx.ord + 1, destLast - 1, 1).getValues()
      .forEach(function(r) { befintliga[(r[0] || '').toString().toLowerCase()] = true; });
  }

  var nyaRader   = [];
  var medGenus   = 0;
  var utanGenus  = 0;
  var dubbletter = 0;

  raw.forEach(function(row) {
    var genus = (row[0] || '').toString().trim().toLowerCase();
    var ord   = (row[1] || '').toString().trim().toLowerCase();
    var cefr  = numCols >= 4 ? (row[3] || '').toString().trim().toUpperCase() : '';
    if (!ord) return;
    if (befintliga[ord]) { dubbletter++; return; }
    befintliga[ord] = true;

    var genusVal = (genus === 'en' || genus === 'ett') ? genus : '';
    var rad = new Array(headers.length).fill('');
    if (idx.ord   >= 0) rad[idx.ord]   = ord;
    if (idx.genus >= 0) rad[idx.genus] = genusVal;
    if (idx.cefr  >= 0 && /^[ABC][12]$/.test(cefr)) rad[idx.cefr] = cefr;
    if (idx.aktiv >= 0) rad[idx.aktiv] = 'ja';

    nyaRader.push(rad);
    if (genusVal) medGenus++; else utanGenus++;
  });

  if (nyaRader.length === 0) {
    SpreadsheetApp.getActive().toast('Inga nya ord att importera.', 'Kelly-import', 6);
    return;
  }

  dest.getRange(dest.getLastRow() + 1, 1, nyaRader.length, headers.length).setValues(nyaRader);
  SpreadsheetApp.getActive().toast(
    nyaRader.length + ' ord importerade  (' + dubbletter + ' dubbletter hoppades över)\n' +
    '✓ Med genus: ' + medGenus + '   ? Saknar genus: ' + utanGenus + '\n' +
    'Kör "Berika med genus via SALDO" för de som saknas.',
    'Kelly-import klar', 12
  );
}

/**
 * Importerar substantiv från Ordbanken i Dokument 1.
 */
function importeraFranOrdbank() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Importera substantiv — steg 1 av 1',
    'Klistra in Google Sheets-ID för dokumentet med Ordbanken.\n' +
    'ID:t finns i URL:en:\ndocs.google.com/spreadsheets/d/[ID]/edit',
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
    var idx     = _getOrdlistaIndex(dest);
    var headers = dest.getRange(1, 1, 1, dest.getLastColumn()).getValues()[0]
      .map(function(h) { return h.toString().toLowerCase(); });

    var befintliga = {};
    var lastDestRow = dest.getLastRow();
    if (lastDestRow >= 2 && idx.ord >= 0) {
      dest.getRange(2, idx.ord + 1, lastDestRow - 1, 1).getValues()
        .forEach(function(r) { befintliga[(r[0] || '').toString().toLowerCase()] = true; });
    }

    var nyaRader = substantiv
      .filter(function(r) { return !befintliga[(r[lemmaIdx] || '').toString().toLowerCase()]; })
      .map(function(r) {
        var rad = new Array(headers.length).fill('');
        if (idx.ord         >= 0) rad[idx.ord]         = (r[lemmaIdx] || '').toString();
        if (idx.deklination >= 0 && deklIdx >= 0) rad[idx.deklination] = (r[deklIdx] || '').toString();
        if (idx.aktiv       >= 0) rad[idx.aktiv]       = 'ja';
        return rad;
      });

    if (nyaRader.length > 0) {
      dest.getRange(dest.getLastRow() + 1, 1, nyaRader.length, headers.length).setValues(nyaRader);
    }
    SpreadsheetApp.getActive().toast(nyaRader.length + ' nya substantiv importerade.', 'Import klar', 10);
  } catch (err) {
    ui.alert('Fel vid import: ' + err.message);
  }
}

// ─── SALDO-berikning ──────────────────────────────────────────────────────────

var SALDO_WS_GAME = 'https://spraakbanken.gu.se/ws/saldo-ws';

/**
 * Slår upp ett ord i SALDO och returnerar paradigm, ordklass och pluralform.
 */
function saldoSlaSuppOrdGame_(word) {
  var url = SALDO_WS_GAME + '/fl/json/' + encodeURIComponent(word);
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { Accept: 'application/json' } });
    if (resp.getResponseCode() !== 200) return { hittad: false };

    var data = JSON.parse(resp.getContentText('UTF-8'));
    if (!Array.isArray(data) || data.length === 0) return { hittad: false };

    var entry    = data[0];
    var paradigm = entry.paradigm || '';
    var prefix   = paradigm.split('_')[0];
    var ordklass = { vb: 'verb', nn: 'substantiv', jj: 'adjektiv' }[prefix] || prefix;

    // Försök extrahera plural indefinit ur formlistan
    var pluralForm = '';
    if (Array.isArray(entry.forms)) {
      var plEntry = null;
      for (var i = 0; i < entry.forms.length; i++) {
        var msd = ((entry.forms[i].msd || entry.forms[i].MSD || '')).toLowerCase();
        if (msd.indexOf('pl') !== -1 && (msd.indexOf('indef') !== -1 || msd.indexOf('ind') !== -1)) {
          plEntry = entry.forms[i];
          break;
        }
      }
      if (plEntry) pluralForm = (plEntry.writtenForm || plEntry.form || '').toLowerCase();
    }

    return { hittad: true, paradigm: paradigm, ordklass: ordklass, plural: pluralForm };
  } catch (e) {
    return { hittad: false };
  }
}

/**
 * Deriverar genus ur SALDO-paradigmkoden.
 * nn_Xu_* → en (utrum)   nn_Xn_* → ett (neutrum)
 */
function tolkaSaldoGenusGame_(paradigm) {
  if (!paradigm || paradigm.indexOf('nn_') !== 0) return '';
  var grp = paradigm.split('_')[1] || '';
  if (grp.charAt(grp.length - 1) === 'u') return 'en';
  if (grp.charAt(grp.length - 1) === 'n') return 'ett';
  return '';
}

/**
 * Fyller i genus och saldo_paradigm för ord utan genus via SALDO-uppslag.
 */
function berikaMedSaldoGenus() {
  var sheet = _getOrdlista();
  if (!sheet) { SpreadsheetApp.getUi().alert('Ordlistan är tom.'); return; }

  var idx     = _getOrdlistaIndex(sheet);
  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var uppdaterade = 0, ejHittade = 0;

  for (var i = 0; i < data.length; i++) {
    var ord      = (data[i][idx.ord]   || '').toString().trim().toLowerCase();
    var harGenus = idx.genus >= 0 ? (data[i][idx.genus] || '').toString().trim() : '';
    if (!ord || harGenus === 'en' || harGenus === 'ett') continue;

    if (i > 0 && i % 15 === 0) SpreadsheetApp.getActive().toast('Bearbetar rad ' + (i + 2) + '…', 'SALDO', 2);
    Utilities.sleep(200);

    var res = saldoSlaSuppOrdGame_(ord);
    if (res.hittad && res.ordklass === 'substantiv') {
      var genus = tolkaSaldoGenusGame_(res.paradigm);
      if (genus) {
        if (idx.genus   >= 0) sheet.getRange(i + 2, idx.genus   + 1).setValue(genus);
        if (idx.paradigm >= 0) sheet.getRange(i + 2, idx.paradigm + 1).setValue(res.paradigm);
        uppdaterade++;
      } else { ejHittade++; }
    } else { ejHittade++; }
  }

  SpreadsheetApp.getActive().toast(
    'Genus ifyllt: ' + uppdaterade + '   Ej hittade / okänt genus: ' + ejHittade,
    'SALDO-berikning klar', 8
  );
}

/**
 * Fyller i plural (pl indef) för ord som saknar plural i Ordlistan.
 * Kör i omgångar (~1 700 ord/körning pga 6 min tidsgräns).
 */
function berikaMedSaldoPlural() {
  var sheet = _getOrdlista();
  if (!sheet) { SpreadsheetApp.getUi().alert('Ordlistan är tom.'); return; }

  var idx = _getOrdlistaIndex(sheet);
  if (idx.plural < 0) {
    SpreadsheetApp.getUi().alert('Ordlistan saknar kolumnen "plural". Kör "Initiera Ordlista-blad" först.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var uppdaterade = 0, ejHittade = 0;

  for (var i = 0; i < data.length; i++) {
    var ord       = (data[i][idx.ord]    || '').toString().trim().toLowerCase();
    var harPlural = (data[i][idx.plural] || '').toString().trim();
    if (!ord || harPlural) continue;

    if (i > 0 && i % 15 === 0) SpreadsheetApp.getActive().toast('Bearbetar rad ' + (i + 2) + '…', 'SALDO-plural', 2);
    Utilities.sleep(200);

    var res = saldoSlaSuppOrdGame_(ord);
    if (res.hittad && res.ordklass === 'substantiv' && res.plural) {
      sheet.getRange(i + 2, idx.plural + 1).setValue(res.plural);
      if (idx.paradigm >= 0 && !(data[i][idx.paradigm] || '').toString().trim()) {
        sheet.getRange(i + 2, idx.paradigm + 1).setValue(res.paradigm);
      }
      uppdaterade++;
    } else { ejHittade++; }
  }

  SpreadsheetApp.getActive().toast(
    'Plural ifyllt: ' + uppdaterade + '   Ej hittade: ' + ejHittade,
    'SALDO-pluralberikning klar', 8
  );
}
