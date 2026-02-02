/**
 * Create a normalized **name key** for a person, robust to:
 * - middle names
 * - swapped first/last name (Japanese style)
 *
 * Steps:
 * 1) Take FIRST WORD of lastName and firstName.
 * 2) Keep only letters, lowercase.
 * 3) If both exist: sort the two tokens alphabetically and join with "|".
 *    If only one exists: use that single token.
 */
function createNameKeyForArchived_(lastName, firstName) {
  function firstWordLettersOnly(s) {
    if (!s) return '';
    var w = String(s).toLowerCase().trim().split(/\s+/)[0]; // first word
    return w.replace(/[^a-z]/g, ''); // letters only
  }

  var t1 = firstWordLettersOnly(lastName);
  var t2 = firstWordLettersOnly(firstName);

  if (t1 && !t2) return t1;
  if (t2 && !t1) return t2;

  if (!t1 || !t2) return '';

  var parts = [t1, t2].sort();
  return parts[0] + '|' + parts[1];
}

/**
 * Create a **person key** using:
 * - Personal ID (preferred when present)
 * - else: name key + birth date
 *
 * NOTE: Personal ID is now in Column Z.
 */
function createPersonKeyForArchived_(personalId, lastName, firstName, birthDate) {
  var pid = (personalId != null) ? String(personalId).trim() : '';
  if (pid) return 'ID||' + pid;

  var nameKey = createNameKeyForArchived_(lastName, firstName);
  if (!nameKey) return '';

  var bd = (birthDate != null && birthDate !== '') ? String(birthDate) : '';
  return nameKey + '||' + bd;
}

/**
 * Build a validation map for one row:
 * For each column, if it has a "list of items" validation, store allowed values.
 */
function buildValidationMap_(sheet, rowIndex, startCol, numCols) {
  var dv = sheet.getRange(rowIndex, startCol, 1, numCols).getDataValidations();
  if (!dv || dv.length === 0) return new Array(numCols).fill(null);

  var validationsRow = dv[0];
  var map = new Array(numCols);

  for (var c = 0; c < numCols; c++) {
    var rule = validationsRow[c];
    if (!rule) {
      map[c] = null;
      continue;
    }

    var critType = rule.getCriteriaType();
    if (critType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var critValues = rule.getCriteriaValues();
      var items = (critValues && critValues[0]) ? critValues[0] : [];
      var allowedSet = {};
      for (var i = 0; i < items.length; i++) {
        var v = items[i];
        if (v != null && v !== '') {
          var norm = String(v).toLowerCase().trim();
          allowedSet[norm] = true;
        }
      }
      map[c] = { type: 'LIST', allowed: allowedSet };
    } else {
      map[c] = null;
    }
  }

  return map;
}

/**
 * Clean a row before writing:
 * - Any string starting with "#" (e.g. #REF!, #VALUE!) → blank.
 * - For columns with list-of-items validation:
 *     if value (lowercased/trimmed) not in allowed list → blank.
 */
function cleanRowForWrite_(row, validationMap) {
  for (var c = 0; c < row.length; c++) {
    var v = row[c];

    if (typeof v === 'string' && v.length > 0 && v.charAt(0) === '#') {
      row[c] = '';
      continue;
    }

    var info = validationMap && validationMap[c] ? validationMap[c] : null;
    if (info && info.type === 'LIST') {
      if (v != null && String(v).trim() !== '') {
        var norm = String(v).toLowerCase().trim();
        if (!info.allowed[norm]) {
          row[c] = '';
        }
      }
    }
  }
  return row;
}

/**
 * Force a row to match the exact number of columns (pad/trim).
 * Prevents: "data has X but range has Y"
 */
function fitRowToColsForArchived_(row, numCols) {
  var out = row.slice(0, numCols);
  while (out.length < numCols) out.push('');
  return out;
}

/**
 * Find the next available (first truly blank) row based on a key column.
 * This avoids appending to the "bottom" when sheets have formatting/validations
 * extending far down.
 */
function findNextAvailableRowForArchived_(sheet, startRow, keyCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;

  var numRows = lastRow - startRow + 1;
  var vals = sheet.getRange(startRow, keyCol, numRows, 1).getValues();

  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    if (v == null || String(v).trim() === '') {
      return startRow + i;
    }
  }
  return lastRow + 1;
}

/**
 * Effective last row finder for Archived:
 * Uses specific columns to detect "real data" rows, avoiding huge getLastRow()
 * caused by formatting/validations far down the sheet.
 */
function findEffectiveLastRowForArchived_(sheet, startRow, cols1Based) {
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow - 1;

  var numRows = lastRow - startRow + 1;
  var maxOffset = -1;

  for (var c = 0; c < cols1Based.length; c++) {
    var col = cols1Based[c];
    var vals = sheet.getRange(startRow, col, numRows, 1).getValues();

    for (var i = vals.length - 1; i >= 0; i--) {
      var v = vals[i][0];
      if (v != null && String(v).trim() !== '') {
        if (i > maxOffset) maxOffset = i;
        break;
      }
    }
  }

  if (maxOffset === -1) return startRow - 1;
  return startRow + maxOffset;
}

/**
 * Find the first blank slot INSIDE archData to fill,
 * where Personal ID (Z) and Last/First (C/D) are all blank.
 */
function findFirstBlankSlotInArchData_(archData, colPid, colLn, colFn) {
  for (var i = 0; i < archData.length; i++) {
    var r = archData[i];
    var pid = (r[colPid] != null) ? String(r[colPid]).trim() : '';
    var ln  = (r[colLn]  != null) ? String(r[colLn]).trim()  : '';
    var fn  = (r[colFn]  != null) ? String(r[colFn]).trim()  : '';
    if (!pid && !ln && !fn) return i;
  }
  return -1;
}

function syncArchivedMembers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var directorySheet = ss.getSheetByName('Directory');
  var archivedSheet = ss.getSheetByName('Archived');
  var configSheet = ss.getSheetByName('Config');

  if (!directorySheet || !archivedSheet || !configSheet) {
    throw new Error('Missing Directory, Archived, or Config sheet.');
  }

  // ===== Layout constants =====
  var DIR_START_ROW   = 4;
  var ARCH_START_ROW  = 4;
  var STATS_START_ROW = 3;

  // 0-based array indexes (A=0, B=1, C=2, ... Z=25)
  var COL_STATUS_A        = 0;   // A (Archived flag)
  var COL_PERSONAL_ID     = 25;  // Z (Personal ID)
  var COL_LAST_NAME       = 2;   // C
  var COL_FIRST_NAME      = 3;   // D
  var COL_BIRTH_DATE      = 6;   // G
  var COL_ASCENDED        = 8;   // I (Directory only)
  var COL_ACTIVITY        = 9;   // J
  var COL_NEW_MEMBER_DATE = 20;  // U

  var lastColDir  = directorySheet.getLastColumn();
  var lastColArch = archivedSheet.getLastColumn();

  var validationMapDir = buildValidationMap_(directorySheet, DIR_START_ROW, 1, lastColDir);

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var SIX_WEEKS_IN_DAYS = 42;

  var lastRowDirInitial = directorySheet.getLastRow();
  var countBeforeDir = Math.max(0, lastRowDirInitial - (DIR_START_ROW - 1));
  var movedToArchived = 0;
  var duplicateMarked = 0;

  // ============================================================
  // 0) Attendance Stats (external file) → nameKey → activity
  // ============================================================
  var attendanceRef = configSheet.getRange('B2').getValue();
  if (!attendanceRef) {
    throw new Error('Config!B2 must contain Attendance sheet URL or ID.');
  }

  var url = String(attendanceRef);
  if (!url.startsWith('http')) {
    url = 'https://docs.google.com/spreadsheets/d/' + url + '/edit';
  }

  var attendanceSs = SpreadsheetApp.openByUrl(url);
  var statsSheet = attendanceSs.getSheetByName('Attendance Stats');
  if (!statsSheet) {
    throw new Error('Attendance Stats tab not found.');
  }

  var statsMap = {};
  var lastRowStats = statsSheet.getLastRow();
  if (lastRowStats >= STATS_START_ROW) {
    var statsValues = statsSheet
      .getRange(STATS_START_ROW, 3, lastRowStats - STATS_START_ROW + 1, 4)
      .getValues(); // C–F

    for (var i = 0; i < statsValues.length; i++) {
      var sRow = statsValues[i];
      var sLast = sRow[0];
      var sFirst = sRow[1];
      var sAct = String(sRow[3] || '').toLowerCase().trim();
      var sNameKey = createNameKeyForArchived_(sLast, sFirst);
      if (sNameKey) statsMap[sNameKey] = sAct;
    }
  }

  // ============================================================
  // 1) Read existing Archived → build personKey map, mark DUP
  // ============================================================
  var lastRowArchExisting = findEffectiveLastRowForArchived_(archivedSheet, ARCH_START_ROW, [1, 3, 4, 26]); // A,C,D,Z
  var archData = [];
  var archPersonIndex = {}; // personKey -> first index

  if (lastRowArchExisting >= ARCH_START_ROW) {
    archData = archivedSheet
      .getRange(ARCH_START_ROW, 1, lastRowArchExisting - ARCH_START_ROW + 1, lastColArch)
      .getValues();

    for (var j = 0; j < archData.length; j++) {
      var rowA = archData[j];

      var pKey = createPersonKeyForArchived_(
        rowA[COL_PERSONAL_ID],
        rowA[COL_LAST_NAME],
        rowA[COL_FIRST_NAME],
        rowA[COL_BIRTH_DATE]
      );
      if (!pKey) continue;

      if (archPersonIndex.hasOwnProperty(pKey)) {
        if (!rowA[COL_STATUS_A]) {
          rowA[COL_STATUS_A] = 'Duplicate Archived Record - ignore this row, keep the first one';
          duplicateMarked++;
        }
      } else {
        archPersonIndex[pKey] = j;
      }
    }
  }

  // ============================================================
  // 2) Directory → archive candidates (NO rewrite; we will delete rows)
  // ============================================================
  var rowsToDelete = [];

  if (lastRowDirInitial >= DIR_START_ROW) {
    var dirData = directorySheet
      .getRange(DIR_START_ROW, 1, lastRowDirInitial - DIR_START_ROW + 1, lastColDir)
      .getValues();

    for (var r = 0; r < dirData.length; r++) {
      var rowD = dirData[r];

      var pid = rowD[COL_PERSONAL_ID];
      var ln  = rowD[COL_LAST_NAME];
      var fn  = rowD[COL_FIRST_NAME];
      var bd  = rowD[COL_BIRTH_DATE];
      var asc = rowD[COL_ASCENDED];
      var act = String(rowD[COL_ACTIVITY] || '').toLowerCase().trim();
      var newMemberDate = rowD[COL_NEW_MEMBER_DATE];

      var isNewMember = false;
      if (newMemberDate) {
        var nmDate = new Date(newMemberDate);
        if (!isNaN(nmDate.getTime())) {
          nmDate.setHours(0, 0, 0, 0);
          var diffMs = today.getTime() - nmDate.getTime();
          var diffDays = diffMs / (1000 * 60 * 60 * 24);
          if (diffDays <= SIX_WEEKS_IN_DAYS) {
            isNewMember = true;
          }
        }
      }

      var statsAct = '';
      var nameKeyDir = createNameKeyForArchived_(ln, fn);
      if (nameKeyDir && statsMap[nameKeyDir]) {
        statsAct = String(statsMap[nameKeyDir]).toLowerCase().trim();
      }

      var hasName =
        (fn && String(fn).trim() !== '') ||
        (ln && String(ln).trim() !== '');

      var ascended = asc != null && String(asc).trim() !== '';
      var shouldArchive = ascended || (
        !isNewMember && (
          act === 'archived' ||
          statsAct === 'archived'
        )
      );

      // Keep existing logic for archived row data cleanliness (Directory formulas are NOT touched)
      var rowClean = cleanRowForWrite_(rowD.slice(), validationMapDir);

      if (!hasName || !shouldArchive) {
        continue;
      }

      var pKey = createPersonKeyForArchived_(pid, ln, fn, bd);

      if (pKey && archPersonIndex.hasOwnProperty(pKey)) {
        if (ascended) {
          var idx = archPersonIndex[pKey];
          archData[idx][COL_STATUS_A] = 'Permanently Archived because ascended';
        }
      } else {
        var newArchRow = rowClean.slice();
        newArchRow[COL_STATUS_A] = ascended
          ? 'Permanently Archived because ascended'
          : (newArchRow[COL_STATUS_A] || '');

        newArchRow = fitRowToColsForArchived_(newArchRow, lastColArch);

        // Put into the next blank slot INSIDE Archived (not bottom)
        var blankIdx = findFirstBlankSlotInArchData_(archData, COL_PERSONAL_ID, COL_LAST_NAME, COL_FIRST_NAME);
        if (blankIdx >= 0) {
          archData[blankIdx] = newArchRow;
          if (pKey) archPersonIndex[pKey] = blankIdx;
        } else {
          archData.push(newArchRow);
          if (pKey) archPersonIndex[pKey] = archData.length - 1;
        }

        movedToArchived++;
      }

      // Mark this Directory row for deletion (preserves formulas/validations)
      rowsToDelete.push(DIR_START_ROW + r);
    }
  }

  // ============================================================
  // 3) Apply changes:
  //    - Delete Directory rows (bottom-up) so formulas are preserved
  //    - Rewrite Archived (as before)
  // ============================================================
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; }); // delete bottom-up
    for (var d = 0; d < rowsToDelete.length; d++) {
      directorySheet.deleteRow(rowsToDelete[d]);
    }
  }

  var existingArchRows = (lastRowArchExisting >= ARCH_START_ROW) ? (lastRowArchExisting - ARCH_START_ROW + 1) : 0;
  var maxRowsToClearArch = Math.max(existingArchRows, archData.length);
  if (maxRowsToClearArch > 0) {
    archivedSheet
      .getRange(ARCH_START_ROW, 1, maxRowsToClearArch, lastColArch)
      .clearContent();
  }

  if (archData.length > 0) {
    for (var z = 0; z < archData.length; z++) {
      archData[z] = fitRowToColsForArchived_(archData[z], lastColArch);
    }

    archivedSheet
      .getRange(ARCH_START_ROW, 1, archData.length, lastColArch)
      .setValues(archData);
  }

  Logger.log("=== ARCHIVE SYNC SUMMARY ===");
  Logger.log("Directory before: " + countBeforeDir);
  Logger.log("Moved to Archived: " + movedToArchived);
  Logger.log("Duplicates marked in Archived: " + duplicateMarked);
  Logger.log("============================");
}

/**
 * Install a trigger to run syncArchivedMembers() every 6 hours.
 * Run this ONCE manually.
 */
function setupTrigger_syncArchivedEverySixHours() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'syncArchivedMembers') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('syncArchivedMembers')
    .timeBased()
    .everyHours(6)
    .create();
}
