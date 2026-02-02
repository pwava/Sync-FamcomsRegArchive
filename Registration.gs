/**
 * Copies the latest form response from "New Registrant" to "Directory".
 * Map columns as specified and format text/date.
 *
 * UPDATE:
 * - If NEW person (not an existing match): issue Personal ID in Directory Column Z
 *   = Config!E8 prefix + 6 alphanumeric (A-Z, 0-9)
 * - Before issuing: check BOTH Directory + Archived Column Z to ensure uniqueness
 * - If EXISTING match: do NOT change Personal ID
 * - Add to the NEXT AVAILABLE ROW (first blank row)
 * - AFTER writing to local Directory:
 *     • Copy the row to ANOTHER Directory (external)
 *     • External sheet link: Config!B3
 *     • External sheet name: "Directory"
 *     • External Column A = Config!E5
 *     • External Column B = LOCAL Column Z (Personal ID)
 *     • External Column C onward = LOCAL Column C onward
 */
function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("New Registrant");
  if (!sheet) return;

  var directorySheet = ss.getSheetByName("Directory");
  if (!directorySheet) return;

  // Get the last submitted row from "New Registrant"
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];

  // --- Helper: Title Case ---
  function toTitleCase_(text) {
    if (!text) return "";
    text = String(text).toLowerCase();
    return text.replace(/\b\w/g, function (c) {
      return c.toUpperCase();
    });
  }

  // --- Helper: matching utilities ---
  function cleanString_(s) {
    if (!s) return "";
    return String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function areAlmostSame_(a, b) {
    if (!a || !b) return false;
    a = String(a); b = String(b);
    var la = a.length, lb = b.length;
    if (Math.abs(la - lb) > 1) return false;
    var i = 0, j = 0, diff = 0;
    while (i < la && j < lb) {
      if (a.charAt(i) === b.charAt(j)) { i++; j++; }
      else {
        diff++; if (diff > 1) return false;
        if (la > lb) i++; else if (lb > la) j++; else { i++; j++; }
      }
    }
    if (i < la || j < lb) diff++;
    return diff <= 1;
  }

  function namesSimilar_(f1, l1, f2, l2) {
    f1 = cleanString_(f1); f2 = cleanString_(f2);
    l1 = cleanString_(l1); l2 = cleanString_(l2);
    if (!l1 || !l2 || l1 !== l2) return false;
    if (!f1 || !f2) return false;
    if (f1 === f2) return true;
    if (f1.indexOf(f2) === 0 || f2.indexOf(f1) === 0) return true;
    return areAlmostSame_(f1, f2);
  }

  function normalizePhone_(p) {
    if (!p) return "";
    return String(p).replace(/\D/g, "");
  }

  function phonesEqual_(a, b) {
    a = normalizePhone_(a);
    b = normalizePhone_(b);
    return a && b && a === b;
  }

  function emailsEqual_(a, b) {
    if (!a || !b) return false;
    return String(a).toLowerCase().trim() === String(b).toLowerCase().trim();
  }

  function getNextAvailableRow_(sheet, startRow, checkCol) {
    var lr = sheet.getLastRow();
    if (lr < startRow) return startRow;
    var vals = sheet.getRange(startRow, checkCol, lr - startRow + 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (!vals[i][0]) return startRow + i;
    }
    return lr + 1;
  }

  // --- Personal ID helpers ---
  function buildExistingIdSet_() {
    var set = {};
    var dLast = directorySheet.getLastRow();
    if (dLast >= 2) {
      var ids = directorySheet.getRange(2, 26, dLast - 1, 1).getValues();
      ids.forEach(function (r) {
        if (r[0]) set[String(r[0]).trim()] = true;
      });
    }
    var archived = ss.getSheetByName("Archived");
    if (archived) {
      var aLast = archived.getLastRow();
      if (aLast >= 2) {
        var aIds = archived.getRange(2, 26, aLast - 1, 1).getValues();
        aIds.forEach(function (r) {
          if (r[0]) set[String(r[0]).trim()] = true;
        });
      }
    }
    return set;
  }

  function randomAlnum6_() {
    var c = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", o = "";
    for (var i = 0; i < 6; i++) o += c.charAt(Math.floor(Math.random() * c.length));
    return o;
  }

  function makeUniquePersonalId_() {
    var cfg = ss.getSheetByName("Config");
    var prefix = cfg ? String(cfg.getRange("E8").getDisplayValue()).trim() : "";
    var used = buildExistingIdSet_();
    while (true) {
      var id = prefix + randomAlnum6_();
      if (!used[id]) return id;
    }
  }

  // --- Values ---
  var timestamp = values[0];
  var firstName = toTitleCase_(values[2]);
  var lastName  = toTitleCase_(values[3]);
  var gender    = toTitleCase_(values[4]);
  var birthdate = values[5];
  var phone     = values[6];
  var email     = values[7];
  var street    = toTitleCase_(values[8]);
  var city      = toTitleCase_(values[9]);
  var state     = values[10];
  var zip       = values[11];
  var invFN     = toTitleCase_(values[12]);
  var invLN     = toTitleCase_(values[13]);
  var invEmail  = values[14];

  // --- Match check ---
  var isExistingMatch = false;
  var existingRow = -1;

  var dLast = directorySheet.getLastRow();
  if (dLast >= 2) {
    var data = directorySheet.getRange(2, 1, dLast - 1, directorySheet.getLastColumn()).getValues();
    var tz = ss.getSpreadsheetTimeZone();
    var regBirthKey = birthdate instanceof Date
      ? Utilities.formatDate(birthdate, tz, "yyyy-MM-dd")
      : "";

    var bestScore = 0;
    var bestIdx = -1;

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!namesSimilar_(firstName, lastName, r[3], r[2])) continue;

      var score = 2;
      if (phonesEqual_(phone, r[14])) score++;
      if (emailsEqual_(email, r[15])) score++;

      if (score > bestScore) {
        bestScore = score;
        bestIdx = i;
      }
    }

    if (bestScore >= 3 && bestIdx >= 0) {
      isExistingMatch = true;
      existingRow = bestIdx + 2;
    }
  }

  var targetRow = isExistingMatch
    ? existingRow
    : getNextAvailableRow_(directorySheet, 2, 3);

  // Personal ID (Column Z)
  if (!isExistingMatch) {
    directorySheet.getRange(targetRow, 26).setValue(makeUniquePersonalId_());
  }

  // Write fields
  directorySheet.getRange(targetRow, 3).setValue(lastName);
  directorySheet.getRange(targetRow, 4).setValue(firstName);
  directorySheet.getRange(targetRow, 5).setValue(gender);
  directorySheet.getRange(targetRow, 7).setValue(birthdate);
  directorySheet.getRange(targetRow, 15).setValue(phone);
  directorySheet.getRange(targetRow, 16).setValue(email);
  directorySheet.getRange(targetRow, 17).setValue(street);
  directorySheet.getRange(targetRow, 18).setValue(city);
  directorySheet.getRange(targetRow, 19).setValue(state);
  directorySheet.getRange(targetRow, 20).setValue(zip);
  directorySheet.getRange(targetRow, 22).setValue(invLN);
  directorySheet.getRange(targetRow, 23).setValue(invFN);
  directorySheet.getRange(targetRow, 24).setValue(invEmail);

  // ==============================
  // COPY TO OTHER DIRECTORY
  // ==============================
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) return;

  var otherUrl = configSheet.getRange("B3").getValue();
  if (!otherUrl) return;

    var otherUrl = configSheet.getRange("B3").getValue();
  if (!otherUrl) return;

  var otherSS;
  if (String(otherUrl).indexOf("http") === 0) {
    otherSS = SpreadsheetApp.openByUrl(otherUrl);
  } else {
    otherSS = SpreadsheetApp.openById(otherUrl);
  }

  var otherDir = otherSS.getSheetByName("Directory");
  if (!otherDir) return;


  var sourceRow = directorySheet.getRange(targetRow, 1, 1, directorySheet.getLastColumn()).getValues()[0];
  var nextOtherRow = getNextAvailableRow_(otherDir, 2, 2);

  // Column A = Config!E5
  otherDir.getRange(nextOtherRow, 1).setValue(configSheet.getRange("E5").getValue());

  // Column B = Personal ID from LOCAL Column Z
  otherDir.getRange(nextOtherRow, 2).setValue(sourceRow[25]);

  // Column C onward = LOCAL Column C onward
  var copyWidth = sourceRow.length - 2;
  if (copyWidth > 0) {
    otherDir
      .getRange(nextOtherRow, 3, 1, copyWidth)
      .setValues([sourceRow.slice(2)]);
  }
}
