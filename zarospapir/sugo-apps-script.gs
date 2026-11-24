/**
 * Lipóti – Zárós papír: súgók („Hol találod meg?”) és szerkesztői belépés
 * ----------------------------------------------------------------------
 * Ez a Google Apps Script egy Google Sheethez tartozik. A Sheetben:
 *   • „Súgók” fül: minden mezőhöz a magyarázat, a kép (a kép a Google Drive-on van), és hogy a dolgozók lássák-e az „i” gombot.
 *   • „Szerkesztők” fül: akik beléphetnek (E-mail, Név, Jelszó).
 *
 * Beállítás (egyszer):
 *   1. Hozz létre egy új Google Sheetet, pl. „Zárós papír – súgók”.
 *   2. Bővítmények → Apps Script. Töröld ki, ami ott van, és másold be ezt az egész fájlt. Mentés.
 *   3. Felül válaszd a „beallitas” függvényt, és kattints a Futtatás gombra. Engedélyezd a hozzáférést.
 *      Ez létrehozza a két fület és a képek mappáját a Drive-on.
 *   4. A „Szerkesztők” fülre írd be a szerkesztőket: E-mail | Név | Jelszó.
 *      A jelszót elég sima szövegként beírni: az első sikeres belépéskor a rendszer titkosítja.
 *   5. Telepítés → Új telepítés → Típus: Webalkalmazás.
 *      Végrehajtás mint: Én. Hozzáférés: Bárki. → Telepítés. Másold ki a „…/exec” végű címet.
 *   6. Ezt a címet kell beírni a zarospapir/index.html fájlba a SUGO_API = '' idézőjelei közé.
 *
 * Kód módosítása után: Telepítés → Telepítések kezelése → ceruza → Verzió: Új verzió → Telepítés.
 * (Így a cím ugyanaz marad.)
 */

var LAP_SUGOK = 'Súgók';
var LAP_SZERK = 'Szerkesztők';
var TOKEN_NAP = 30;            // ennyi napig marad bejelentkezve a szerkesztő
var MAX_HIBA = 5;              // ennyi rossz jelszó után 15 percre zár
var MAX_KEP_BAJT = 4 * 1024 * 1024;

/* ===================== Belépési pont ===================== */

function doGet() {
  return valasz({ ok: true, data: { nev: 'Lipóti – Zárós papír súgók', ido: new Date().toISOString() } });
}

function doPost(e) {
  try {
    var keres = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = String(keres.action || '');
    var p = keres.payload || {};
    var kulcs = String(keres.key || '');
    switch (action) {
      case 'sugok':   return ok(sugokOlvas());
      case 'kep':     return ok(kepOlvas(p.id));
      case 'belepes': return ok(belepes(p.email, p.jelszo));
      case 'en':      return ok(szerkesztoKell(kulcs));
      case 'kilepes': tokenTorol(kulcs); return ok({});
      case 'mentes':  return ok(mentes(szerkesztoKell(kulcs), p));
      default:        return hiba('Ismeretlen kérés.', 'BAD');
    }
  } catch (err) {
    if (err && err.kod) return hiba(err.message, err.kod);
    console.error(err && err.stack || err);
    return hiba('Hiba történt a szerveren. Próbáld újra pár perc múlva.', 'SERVER');
  }
}

function valasz(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ok(data) { return valasz({ ok: true, data: data }); }
function hiba(szoveg, kod) { return valasz({ ok: false, error: szoveg, code: kod || 'ERR' }); }
function dob(szoveg, kod) { var e = new Error(szoveg); e.kod = kod; throw e; }

/* ===================== Súgók ===================== */

function sugokOlvas() {
  var cache = CacheService.getScriptCache();
  var c = cache.get('sugok');
  if (c) return JSON.parse(c);
  var lap = lapKell(LAP_SUGOK);
  var adat = lap.getDataRange().getValues();
  var ki = {};
  for (var i = 1; i < adat.length; i++) {
    var k = String(adat[i][0] || '').trim();
    if (!k) continue;
    var l = String(adat[i][6] || '').trim().toLowerCase();
    ki[k] = { szoveg: String(adat[i][2] || ''), kepId: String(adat[i][3] || '') || null, lathato: l === '' ? null : l === 'igen' };
  }
  cache.put('sugok', JSON.stringify(ki), 600);
  return ki;
}

function kepOlvas(id) {
  id = String(id || '');
  if (!/^[A-Za-z0-9_-]{10,}$/.test(id)) dob('Hibás kép.', 'BAD');
  var sugok = sugokOlvas(), van = false;
  for (var k in sugok) if (sugok[k].kepId === id) { van = true; break; }
  if (!van) dob('Ez a kép már nincs meg.', 'NOTFOUND');
  var blob = DriveApp.getFileById(id).getBlob();
  return { mime: blob.getContentType() || 'image/jpeg', base64: Utilities.base64Encode(blob.getBytes()) };
}

function mentes(szerk, p) {
  var kulcs = String(p.kulcs || '').trim();
  if (!/^[A-Za-z0-9_-]{1,40}$/.test(kulcs)) dob('Hibás mező.', 'BAD');
  var szoveg = String(p.szoveg || '').slice(0, 5000);
  var cim = String(p.cim || kulcs).slice(0, 200);

  var zar = LockService.getScriptLock();
  zar.waitLock(20000);
  try {
    var lap = lapKell(LAP_SUGOK);
    var adat = lap.getDataRange().getValues();
    var sor = -1;
    for (var i = 1; i < adat.length; i++) if (String(adat[i][0]).trim() === kulcs) { sor = i + 1; break; }
    var regiKep = sor > 0 ? String(adat[sor - 1][3] || '') : '';
    var kepId = p.kepMarad ? regiKep : '';

    if (p.kep && p.kep.base64) {
      var mime = String(p.kep.mime || 'image/jpeg');
      if (!/^image\/(jpeg|png|webp|gif)$/.test(mime)) dob('Csak JPG vagy PNG képet tölthetsz fel.', 'BAD');
      var bajtok = Utilities.base64Decode(p.kep.base64);
      if (bajtok.length > MAX_KEP_BAJT) dob('A kép túl nagy. Válassz kisebb képet.', 'BAD');
      var fajl = kepMappa().createFile(Utilities.newBlob(bajtok, mime, kulcs + '_' + Date.now() + (mime === 'image/png' ? '.png' : '.jpg')));
      kepId = fajl.getId();
    }
    if (regiKep && regiKep !== kepId) { try { DriveApp.getFileById(regiKep).setTrashed(true); } catch (e) {} }

    var most = Utilities.formatDate(new Date(), 'Europe/Budapest', 'yyyy.MM.dd. HH:mm');
    var lathato = p.lathato === true ? 'igen' : (p.lathato === false ? 'nem' : '');
    var ertek = [[kulcs, cim, szoveg, kepId, most, szerk.nev || szerk.email, lathato]];
    if (sor > 0) lap.getRange(sor, 1, 1, 7).setValues(ertek);
    else { lap.appendRow(ertek[0]); formaz(lap); }
    CacheService.getScriptCache().remove('sugok');
    return { szoveg: szoveg, kepId: kepId || null, lathato: lathato === '' ? null : lathato === 'igen' };
  } finally {
    zar.releaseLock();
  }
}

/* ===================== Belépés ===================== */

function belepes(email, jelszo) {
  email = String(email || '').trim().toLowerCase();
  jelszo = String(jelszo || '');
  if (!email || !jelszo) dob('Írd be az e-mail-címedet és a jelszavadat.', 'BAD');

  var cache = CacheService.getScriptCache();
  var hibaKulcs = 'hiba_' + email;
  var hibak = Number(cache.get(hibaKulcs) || 0);
  if (hibak >= MAX_HIBA) dob('Túl sok rossz próbálkozás. Várj 15 percet, és próbáld újra.', 'LOCKED');

  var lap = lapKell(LAP_SZERK);
  var adat = lap.getDataRange().getValues();
  for (var i = 1; i < adat.length; i++) {
    if (String(adat[i][0] || '').trim().toLowerCase() !== email) continue;
    var tarolt = String(adat[i][2] || '');
    var jo = false;
    if (tarolt.indexOf('sha256:') === 0) {
      var r = tarolt.split(':');               // sha256:só:hash
      jo = r.length === 3 && hash(r[1], jelszo) === r[2];
    } else if (tarolt && tarolt === jelszo) {
      jo = true;                               // első belépés: sima jelszó → titkosítjuk
      var so = Utilities.getUuid();
      lap.getRange(i + 1, 3).setValue('sha256:' + so + ':' + hash(so, jelszo));
    }
    if (jo) {
      cache.remove(hibaKulcs);
      tokenTakarit();
      var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
      PropertiesService.getScriptProperties().setProperty('t_' + token,
        JSON.stringify({ email: email, nev: String(adat[i][1] || ''), lejar: Date.now() + TOKEN_NAP * 864e5 }));
      return { token: token, email: email, nev: String(adat[i][1] || '') };
    }
    break;
  }
  cache.put(hibaKulcs, String(hibak + 1), 900);
  dob('Hibás e-mail-cím vagy jelszó. Ellenőrizd, és próbáld újra.', 'AUTH');
}

function szerkesztoKell(token) {
  if (!token) dob('Ehhez be kell lépned.', 'AUTH');
  var p = PropertiesService.getScriptProperties();
  var adat = p.getProperty('t_' + token);
  if (!adat) dob('Lejárt a belépésed. Lépj be újra.', 'AUTH');
  var t = JSON.parse(adat);
  if (t.lejar < Date.now()) { p.deleteProperty('t_' + token); dob('Lejárt a belépésed. Lépj be újra.', 'AUTH'); }
  // aki már nincs a Szerkesztők fülön, az nem szerkeszthet
  var lista = lapKell(LAP_SZERK).getDataRange().getValues().slice(1).map(function (r) { return String(r[0] || '').trim().toLowerCase(); });
  if (lista.indexOf(t.email) < 0) { p.deleteProperty('t_' + token); dob('Már nincs szerkesztői jogod.', 'AUTH'); }
  return t;
}

function tokenTorol(token) { if (token) PropertiesService.getScriptProperties().deleteProperty('t_' + token); }

function tokenTakarit() {
  var p = PropertiesService.getScriptProperties(), osszes = p.getProperties(), most = Date.now();
  Object.keys(osszes).forEach(function (k) {
    if (k.indexOf('t_') !== 0) return;
    try { if (JSON.parse(osszes[k]).lejar < most) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); }
  });
}

function hash(so, jelszo) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, so + '|' + jelszo, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

/* ===================== Beállítás és formázás ===================== */

function beallitas() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(LAP_SUGOK) || ss.insertSheet(LAP_SUGOK);
  if (s.getLastRow() === 0) s.appendRow(['Kulcs', 'Mező', 'Magyarázat', 'Kép (Drive fájl)', 'Frissítve', 'Frissítette', 'Látható a dolgozóknak (igen/nem)']);
  else if (!s.getRange(1, 7).getValue()) s.getRange(1, 7).setValue('Látható a dolgozóknak (igen/nem)');
  var z = ss.getSheetByName(LAP_SZERK) || ss.insertSheet(LAP_SZERK);
  if (z.getLastRow() === 0) z.appendRow(['E-mail', 'Név', 'Jelszó']);
  [s, z].forEach(formaz);
  s.setColumnWidth(1, 110); s.setColumnWidth(2, 260); s.setColumnWidth(3, 480); s.setColumnWidth(4, 200); s.setColumnWidth(5, 140); s.setColumnWidth(6, 160); s.setColumnWidth(7, 170);
  s.getRange('C:C').setWrap(true);
  z.setColumnWidth(1, 260); z.setColumnWidth(2, 200); z.setColumnWidth(3, 320);
  var lap1 = ss.getSheetByName('Munkalap1') || ss.getSheetByName('Sheet1');
  if (lap1 && lap1.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(lap1);
  kepMappa();
  Logger.log('Kész. Töltsd ki a Szerkesztők fület, majd telepítsd webalkalmazásként.');
}

/** A szabálykönyv 6. pontja: barna fejléc, középre, rögzített első sor, vékony keretek a kitöltött részen. */
function formaz(lap) {
  var oszlop = Math.max(lap.getLastColumn(), 1), sor = Math.max(lap.getLastRow(), 1);
  lap.setFrozenRows(1);
  lap.setRowHeight(1, 36);
  lap.getRange(1, 1, 1, oszlop).setBackground('#823D02').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  lap.getRange(1, 1, sor, oszlop).setBorder(true, true, true, true, true, true, '#C9B79C', SpreadsheetApp.BorderStyle.SOLID);
}

function kepMappa() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('kepMappa');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var m = DriveApp.createFolder('Zárós papír – súgóképek');
  p.setProperty('kepMappa', m.getId());
  return m;
}

function lapKell(nev) {
  var lap = SpreadsheetApp.getActive().getSheetByName(nev);
  if (!lap) dob('A „' + nev + '” fül hiányzik. Futtasd a „beallitas” függvényt.', 'SERVER');
  return lap;
}
