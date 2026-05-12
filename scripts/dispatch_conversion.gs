/**
 * 月別横持ち配車表 → 10_配車予定 縦持ち変換（GAS 第1版）
 *
 * @fileoverview 配車予定表を読み取り、システム連携用シートへ UPSERT する。
 * @see docs/dispatch_conversion_design.md
 * @see docs/dispatch_conversion_gas_spec.md
 * @see docs/schema_master.md（B. 10_配車予定）
 */

/** @type {string} 出力先シート名（実体と一致させる） */
var TARGET_SCHEDULE_SHEET_NAME = '10_配車予定';

/** @type {number} 案件列見出しの行番号（1起点・実ファイルに合わせて変更） */
var HEADER_ROW = 2;

/** @type {number} データ走査開始行（1起点・実ファイルに合わせて変更） */
var START_ROW = 3;

/** @type {number} A列＝日付 */
var DATE_COL = 1;

/** @type {number} B列＝曜日 */
var WEEKDAY_COL = 2;

/** @type {number} C列＝区分 */
var TYPE_COL = 3;

/** @type {number} 案件開始列（D列） */
var FIRST_JOB_COL = 4;

/**
 * 案件列の右端を決めるヘッダー見出し名（この列の1つ左までが案件列）。
 * 「担当者」列およびその右（出勤日数等）は変換対象外。
 */
var STOP_HEADER_NAME = '担当者';

/** @type {string} 担当者行の区分値 */
var TYPE_DRIVER = '担当者';

/** @type {string} 車両行の区分値 */
var TYPE_VEHICLE = '車両';

/** @type {string} ステータス初期値 */
var INITIAL_STATUS = '確定';

/** @type {string} マスタ照合：未実施 */
var MASTER_PENDING = '未照合';

/** @type {string} マスタ照合：要人手確認 */
var MASTER_NEEDS_REVIEW = '要確認';

/** @type {string} メニューラベル */
var MENU_DISPATCH = '配車変換';

/** @type {string} メニュー項目 */
var MENU_CONVERT_CURRENT = '現在のシートを10_配車予定へ変換';

/** 10_配車予定に必須とするヘッダー（1行目に存在すること） */
var REQUIRED_SCHEDULE_HEADERS = [
  '運行予定ID',
  '予定日',
  '案件名_入力値',
  '運転者名_入力値',
  '車両呼称_入力値',
  '作成元シート名',
  '作成元セル',
  'マスタ照合状態',
  '更新日時',
];

/**
 * スプレッドシート起動時：カスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_DISPATCH)
    .addItem(MENU_CONVERT_CURRENT, 'convertCurrentDispatchSheetToSchedule')
    .addItem('決定案件を10_配車予定へ反映', 'reflectDecidedCandidatesToSchedule')
    .addItem('未配車案件の割当状態を更新', 'updateAssignedDispatchStatus')
    .addToUi();
}

/**
 * アクティブシート（月別配車表）を読み、10_配車予定へ変換する。
 */
function convertCurrentDispatchSheetToSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  try {
    if (sourceSheet.getName() === TARGET_SCHEDULE_SHEET_NAME) {
      throw new Error(
        'アクティブシートは変換対象の月別配車表にしてください（「' +
          TARGET_SCHEDULE_SHEET_NAME +
          '」は変換先のため実行できません）。'
      );
    }

    var scheduleSheet = ss.getSheetByName(TARGET_SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) {
      throw new Error('シート「' + TARGET_SCHEDULE_SHEET_NAME + '」が見つかりません。');
    }

    var headerMap = getSheetHeaderMap_(scheduleSheet);
    validateRequiredHeaders_(headerMap);

    var values = readDispatchSheet_(sourceSheet);
    var jobNames = getJobNames_(values);

    var stats = { skippedBothEmpty: 0, skippedNoDate: 0 };
    var rows = buildScheduleRows_(sourceSheet, values, jobNames, stats);
    var existingMap = findExistingScheduleMap_(scheduleSheet);

    var result = upsertScheduleRows_(scheduleSheet, rows, existingMap);

    var needsReview = rows.filter(function (r) {
      return r['マスタ照合状態'] === MASTER_NEEDS_REVIEW;
    }).length;

    var msg =
      '変換完了：\n' +
      '追加 ' +
      result.added +
      '件\n' +
      '更新 ' +
      result.updated +
      '件\n' +
      '要確認 ' +
      needsReview +
      '件\n' +
      'スキップ ' +
      (stats.skippedBothEmpty + stats.skippedNoDate) +
      '件';

    Logger.log(msg);
    Logger.log(
      '詳細 skipempty=' +
        stats.skippedBothEmpty +
        ' skipNoDate=' +
        stats.skippedNoDate +
        ' upsertSkipped=' +
        result.skipped
    );

    ui.alert(msg);
  } catch (e) {
    Logger.log(e.stack || String(e.message));
    ui.alert('エラー: ' + (e.message || String(e)));
  }
}

/**
 * 対象シートのすべてのデータを読み取る。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {!Array<!Array<*>>}
 */
function readDispatchSheet_(sheet) {
  return sheet.getDataRange().getValues();
}

/**
 * HEADER_ROW で「STOP_HEADER_NAME」列の直前までを案件列とし、その見出しから案件名一覧を構築する。
 * @param {!Array<!Array<*>>} values
 * @return {!Array<{colIndex: number, a1Column: string, jobName: string}>}
 */
function getJobNames_(values) {
  var lastJobCol = getLastJobCol_(values);
  var header = values[HEADER_ROW - 1];
  if (!header) throw new Error('ヘッダー行が見つかりません: ' + HEADER_ROW);

  var jobNames = [];

  for (var col = FIRST_JOB_COL; col <= lastJobCol; col++) {
    var jobName = normalizeCellValue_(header[col - 1]);
    if (!jobName) continue;
    jobNames.push({
      colIndex: col,
      a1Column: columnToLetters_(col),
      jobName: jobName,
    });
  }

  if (jobNames.length === 0) {
    throw new Error(
      '案件ヘッダを取得できませんでした（HEADER_ROW=' + HEADER_ROW + ' を確認してください）。'
    );
  }

  Logger.log('案件列範囲: ' + FIRST_JOB_COL + '〜' + lastJobCol);
  Logger.log('案件列数: ' + jobNames.length);

  return jobNames;
}

/**
 * ヘッダー行から指定名と一致する列番号を返す（1起点）。見つからない場合は 0。
 * @param {!Array<!Array<*>>} values
 * @param {number} headerRow 1起点
 * @param {string} headerName 正規化後と比較する見出し名
 * @param {number=} startCol 検索開始列（1起点）。省略時は 1
 * @return {number}
 */
function findHeaderColumnByName_(values, headerRow, headerName, startCol) {
  var row = values[headerRow - 1];
  if (!row) {
    throw new Error('ヘッダー行が見つかりません: ' + headerRow);
  }
  var fromCol = startCol == null || startCol < 1 ? 1 : startCol;
  for (var col = fromCol; col <= row.length; col++) {
    var value = normalizeCellValue_(row[col - 1]);
    if (value === headerName) {
      return col;
    }
  }
  return 0;
}

/**
 * 案件列の最終列（1起点）。「担当者」ヘッダ列の1つ左。
 * @param {!Array<!Array<*>>} values
 * @return {number}
 */
function getLastJobCol_(values) {
  Logger.log(`HEADER_ROW=${HEADER_ROW}`);
  Logger.log(`STOP_HEADER_NAME=${STOP_HEADER_NAME}`);

  var stopCol = findHeaderColumnByName_(
    values,
    HEADER_ROW,
    STOP_HEADER_NAME,
    FIRST_JOB_COL
  );
  if (!stopCol) {
    throw new Error(
      '案件列の終了位置を判定できませんでした。ヘッダー行のD列以降に「' +
        STOP_HEADER_NAME +
        '」があるか確認してください。'
    );
  }

  var lastJobCol = stopCol - 1;

  Logger.log(`stopCol=${stopCol}`);
  Logger.log(`lastJobCol=${lastJobCol}`);

  if (lastJobCol < FIRST_JOB_COL) {
    throw new Error(
      '案件列の範囲が不正です。FIRST_JOB_COL=' +
        FIRST_JOB_COL +
        ', stopCol=' +
        stopCol +
        ', lastJobCol=' +
        lastJobCol
    );
  }

  return lastJobCol;
}

/**
 * 親スプレッドシートのファイル名から年（例: 2026）を取得する。
 * @param {string} spreadsheetName
 * @return {number}
 */
function getYearFromSpreadsheetName_(spreadsheetName) {
  var text = normalizeCellValue_(String(spreadsheetName));
  var match = text.match(/(20\d{2})年?/);
  if (!match) {
    throw new Error('スプレッドシート名から年を取得できませんでした: ' + spreadsheetName);
  }
  return Number(match[1]);
}

/**
 * 月別シート名から月を取得する（例: 「4月度」→ 4）。
 * @param {string} sheetName
 * @return {number}
 */
function getMonthFromSheetName_(sheetName) {
  var text = normalizeCellValue_(String(sheetName));
  var match = text.match(/(\d{1,2})月/);
  if (!match) {
    throw new Error('シート名から月を取得できませんでした: ' + sheetName);
  }
  var month = Number(match[1]);
  if (month < 1 || month > 12) {
    throw new Error('シート名から取得した月が不正です: ' + sheetName);
  }
  return month;
}

/**
 * A列の日番号（1〜31）を取得。Excelシリアルのような大きな数値は日として扱わない。
 * @param {*} value
 * @return {?number}
 */
function getDayNumber_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getDate();
  }
  if (typeof value === 'number' && !isNaN(value)) {
    var n = Math.floor(Number(value));
    if (Number(value) === n && n >= 1 && n <= 31) return n;
    return null;
  }
  var text = normalizeCellValue_(value);
  if (!text) return null;
  if (!/^\d{1,2}$/.test(text)) return null;
  var day = parseInt(text, 10);
  if (day < 1 || day > 31) return null;
  return day;
}

/**
 * 年月日から「その暦日」の Date を返す（不正な組み合わせは null）。
 * **`10_配車予定` の予定日出力には使わず**、`buildScheduleDateText_` と曜日算出用 Date の検証など補助用。
 * @param {number} year
 * @param {number} month
 * @param {number} day
 * @return {?Date}
 */
function buildScheduleDate_(year, month, day) {
  if (!year || !month || !day) return null;
  var d = new Date(year, month - 1, day, 12, 0, 0);
  if (
    isNaN(d.getTime()) ||
    d.getFullYear() !== year ||
    d.getMonth() !== month - 1 ||
    d.getDate() !== day
  ) {
    return null;
  }
  return d;
}

/**
 * 10_配車予定へ出力する予定日文字列 yyyy/MM/dd（TZ により Date 書き込みで前日表示になることを避ける）。
 * @param {number} year
 * @param {number} month
 * @param {number} day
 * @return {string}
 */
function buildScheduleDateText_(year, month, day) {
  if (!year || !month || !day) return '';
  var yyyy = String(year);
  var mm = String(month).padStart(2, '0');
  var dd = String(day).padStart(2, '0');
  return yyyy + '/' + mm + '/' + dd;
}

/**
 * 予定日から曜日表示（例: 「水曜日」）を返す。第1版の曜日出力の正。
 * 算出用の **正午ローカル Date** と整合（年月日のみの暦日上の曜日）。
 * @param {Date} dateValue
 * @return {string}
 */
function getWeekdayText_(dateValue) {
  if (!(dateValue instanceof Date) || isNaN(dateValue.getTime())) return '';
  var weekdays = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];
  return weekdays[dateValue.getDay()];
}

/**
 * 担当者行＋車両行のセットから 10_配車予定 向けオブジェクトを生成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {!Array<!Array<*>>} values
 * @param {!Array<{colIndex: number, a1Column: string, jobName: string}>} jobNames
 * @param {{skippedBothEmpty: number, skippedNoDate: number}=} statistics （省略時は内部初期化のみ）
 * @return {!Array<!Object<string, *>>}
 */
function buildScheduleRows_(sheet, values, jobNames, statistics) {
  statistics = statistics || {};
  if (typeof statistics.skippedBothEmpty !== 'number') statistics.skippedBothEmpty = 0;
  if (typeof statistics.skippedNoDate !== 'number') statistics.skippedNoDate = 0;

  var sheetName = sheet.getName();
  var spreadsheetName = sheet.getParent().getName();
  var year = getYearFromSpreadsheetName_(spreadsheetName);
  var month = getMonthFromSheetName_(sheetName);

  var out = [];

  var typeColIdx = TYPE_COL - 1;
  var startIdx = START_ROW - 1;

  for (var i = startIdx; i < values.length; i++) {
    var row = values[i];
    if (!row) continue;

    var typeCell = normalizeCellValue_(row[typeColIdx]);
    if (typeCell !== TYPE_DRIVER) continue;

    var driverRowIdx = i;
    var driverRow = values[driverRowIdx];
    var vehicleRow = values[i + 1];
    var vehicleRowMissing =
      !vehicleRow || normalizeCellValue_(vehicleRow[typeColIdx]) !== TYPE_VEHICLE;

    var day = getDayNumber_(driverRow[DATE_COL - 1]);
    if (!day) {
      statistics.skippedNoDate++;
      continue;
    }
    var scheduleDateForCalc = new Date(year, month - 1, day, 12, 0, 0);
    if (
      isNaN(scheduleDateForCalc.getTime()) ||
      scheduleDateForCalc.getFullYear() !== year ||
      scheduleDateForCalc.getMonth() !== month - 1 ||
      scheduleDateForCalc.getDate() !== day
    ) {
      statistics.skippedNoDate++;
      continue;
    }
    var scheduleDateText = buildScheduleDateText_(year, month, day);
    if (!scheduleDateText) {
      statistics.skippedNoDate++;
      continue;
    }
    var weekday = getWeekdayText_(scheduleDateForCalc);

    for (var j = 0; j < jobNames.length; j++) {
      var job = jobNames[j];
      var ci = job.colIndex - 1;
      var drv = normalizeCellValue_(driverRow[ci]);
      var veh = vehicleRowMissing ? '' : normalizeCellValue_(vehicleRow[ci]);

      if (!drv && !veh) {
        statistics.skippedBothEmpty++;
        continue;
      }

      var needReviewPartial = !!(drv !== '' && veh === '') || !!(drv === '' && veh !== '');
      var needReviewLayout = !!vehicleRowMissing;
      var matchState = MASTER_PENDING;
      if (needReviewPartial || needReviewLayout) matchState = MASTER_NEEDS_REVIEW;

      var driverA1 = toA1_(driverRowIdx + 1, job.colIndex);

      var memoParts = [];
      if (vehicleRowMissing) memoParts.push('直下に車両行がありません（要確認）。');
      if (needReviewPartial) memoParts.push('担当者または車両の片方のみ入力（要確認）。');
      var memo = memoParts.join(' ');

      /** @type {!Object<string, *>} */
      var rec = {};

      rec['運行予定ID'] = '';
      rec['運行予定番号'] = '';
      rec['予定日'] = scheduleDateText;
      rec['曜日'] = weekday;
      rec['案件ID'] = '';
      rec['案件名_入力値'] = job.jobName;
      rec['運転者ID'] = '';
      rec['運転者名_入力値'] = drv;
      rec['車両ID'] = '';
      rec['車両呼称_入力値'] = veh;

      rec['ステータス'] = INITIAL_STATUS;
      rec['出発予定時刻'] = '';
      rec['到着予定時刻'] = '';
      rec['出発地'] = '';
      rec['目的地'] = '';

      // schema_master 側の項目名および将来列用エイリアス（ヘッダが無ければ出力時無視）

      rec['作業内容'] = '';
      rec['備考'] = '';
      rec['摘要・ルート等'] = '';
      rec['カレンダー表示名'] = '';
      rec['カレンダーイベントID'] = '';
      rec['最終同期日時'] = '';
      rec['LINE送信済みフラグ'] = false;
      rec['取引先ID'] = '';

      rec['作成元シート名'] = sheetName;
      rec['作成元セル'] = driverA1;
      rec['マスタ照合状態'] = matchState;
      rec['照合エラーメモ'] = memo;
      rec['作成日時'] = '';
      rec['更新日時'] = '';

      out.push(rec);
    }
  }

  return out;
}

/**
 * 既存の 10_配車予定 を複合キーでインデックス化する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} scheduleSheet
 * @return {!Object<string, {rowNumber: number, rowValues: !Array<*>, rowObject: !Object}>}
 */
function findExistingScheduleMap_(scheduleSheet) {
  var data = scheduleSheet.getDataRange().getValues();
  if (data.length < 2) return {};

  var headers = data[0];
  /** @type {!Object<string, number>} */
  var colByName = {};
  for (var h = 0; h < headers.length; h++) {
    var hn = normalizeCellValue_(headers[h]);
    if (hn) colByName[hn] = h;
  }

  var map = {};
  var numCols = headers.length;

  for (var r = 1; r < data.length; r++) {
    var sheetRowNum = r + 1;
    var rowVals = data[r];
    /** @type {!Object<string, *>} */
    var robj = {};

    var names = Object.keys(colByName);
    for (var k = 0; k < names.length; k++) {
      var name = names[k];
      robj[name] = rowVals[colByName[name]];
    }

    var key = buildSourceKey_(robj);
    if (!key) continue;

    map[key] = {
      rowNumber: sheetRowNum,
      rowValues: rowVals.concat().slice(0, numCols),
      rowObject: robj,
    };
  }

  return map;
}

/**
 * setValues 用に1行を物理列数へ整える。無効時は null。
 * @param {*} arr
 * @param {number} numColsPhysical
 * @return {?Array}
 */
function sanitizeLineForSet_(arr, numColsPhysical) {
  if (!Array.isArray(arr)) return null;
  var line = [];
  for (var c = 0; c < numColsPhysical; c++) {
    line.push(c < arr.length ? arr[c] : '');
  }
  return line;
}

/**
 * 物理列上ですべて空白相当か（空セル／null／undefined）。
 * @param {!Array<*>} line
 * @return {boolean}
 */
function isBlankPhysicalLine_(line) {
  for (var i = 0; i < line.length; i++) {
    var v = line[i];
    if (v === null || v === undefined) continue;
    if (v === '') continue;
    if (v instanceof Date) {
      if (!isNaN(v.getTime())) return false;
      continue;
    }
    return false;
  }
  return true;
}

/**
 * 追記候補行から setValues に渡せるものだけ残す。
 * @param {!Array} rows2d 二次元配列（undefined 混入可）
 * @param {number} numColsPhysical
 * @return {!Array<!Array<*>>}
 */
function filterAppendLinesForSet_(rows2d, numColsPhysical) {
  var out = [];
  if (!rows2d || !rows2d.length) return out;
  for (var i = 0; i < rows2d.length; i++) {
    var row = rows2d[i];
    if (row === null || row === undefined) continue;
    var sanitized = sanitizeLineForSet_(row, numColsPhysical);
    if (!sanitized || isBlankPhysicalLine_(sanitized)) continue;
    out.push(sanitized);
  }
  return out;
}

/**
 * UPSERT で 10_配車予定 を更新または追記する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} scheduleSheet
 * @param {!Array<!Object<string, *>>} rows
 * @param {!Object<string, *>} existingMap buildSourceKey_ → 既存情報
 * @return {{added: number, updated: number, skipped: number, needsReview: number}}
 */
function upsertScheduleRows_(scheduleSheet, rows, existingMap) {
  existingMap = existingMap || {};

  var lastColHdr = scheduleSheet.getLastColumn();
  if (lastColHdr < 1) throw new Error('10_配車予定のヘッダーがありません。');

  var headerRowVals = scheduleSheet.getRange(1, 1, 1, lastColHdr).getValues()[0];

  /** 物理列順序をそのまま使う（空ヘッダー列は出力を空とする）。 */
  var numColsPhysical = headerRowVals.length;

  /** 変換0件でもエラーにせず打ち切りログのみ出す */
  if (!rows || rows.length === 0) {
    Logger.log('追加対象行数: 0');
    Logger.log('更新対象行数: 0');
    Logger.log('実際に setValues する行数: 0');
    Logger.log(
      'UPSERT 省略（入力0件） skipped=0 needsReview(in rows)=0'
    );
    return { added: 0, updated: 0, skipped: 0, needsReview: 0 };
  }

  var now = new Date();

  var skipped = 0;

  /** @type {!Array<{rn: number, arr: !Array<?>}>} */
  var rowUpdatesToApply = [];
  /** @type {!Array<!Array<*>>} */
  var appendQueue = [];

  /** 論理カウント（キューへ入れた追加・更新の候補数） */
  var appendQueued = 0;
  var updateQueued = 0;

  for (var i = 0; i < rows.length; i++) {
    var src = rows[i];

    /** @type {!Object<string, *>} */
    var merged = {};

    var k = buildSourceKey_(src);
    if (!k) {
      skipped++;
      continue;
    }

    var ex = existingMap[k];
    if (ex && ex.rowObject) {
      var base = ex.rowObject;
      var names = Object.keys(base);
      for (var n = 0; n < names.length; n++) {
        merged[names[n]] = base[names[n]];
      }
    }

    var writeNames = Object.keys(src);
    for (var w = 0; w < writeNames.length; w++) {
      var wn = writeNames[w];
      merged[wn] = src[wn];
    }

    if (ex) {
      merged['運行予定ID'] = ex.rowObject['運行予定ID'];

      merged['カレンダーイベントID'] =
        ex.rowObject['カレンダーイベントID'] != null &&
        String(ex.rowObject['カレンダーイベントID']) !== ''
          ? ex.rowObject['カレンダーイベントID']
          : merged['カレンダーイベントID'];

      merged['最終同期日時'] =
        ex.rowObject['最終同期日時'] != null &&
        String(ex.rowObject['最終同期日時']) !== ''
          ? ex.rowObject['最終同期日時']
          : merged['最終同期日時'];

      merged['LINE送信済みフラグ'] = ex.rowObject['LINE送信済みフラグ'];

      if (ex.rowObject['作成日時'] != null && String(ex.rowObject['作成日時']) !== '') {
        merged['作成日時'] = ex.rowObject['作成日時'];
      }
    } else {
      merged['運行予定ID'] = createScheduleId_();
      merged['作成日時'] = now;
      merged['LINE送信済みフラグ'] = false;
    }

    merged['更新日時'] = now;

    var line = buildOutputLine_(headerRowVals, merged);

    if (ex) {
      rowUpdatesToApply.push({ rn: ex.rowNumber, arr: line });
      updateQueued++;
    } else {
      appendQueue.push(line);
      appendQueued++;
    }
  }

  /** 書き込み対象のみ（無効・空行を除外） */
  var appendValues = filterAppendLinesForSet_(appendQueue, numColsPhysical);

  /** @type {!Array<{rn: number, line: !Array<*>}>} */
  var updateWrites = [];
  for (var u0 = 0; u0 < rowUpdatesToApply.length; u0++) {
    var patch0 = rowUpdatesToApply[u0];
    var lineU = sanitizeLineForSet_(patch0.arr, numColsPhysical);
    if (!lineU) continue;
    updateWrites.push({ rn: patch0.rn, line: lineU });
  }

  var appendedSkipped = appendQueued - appendValues.length;
  if (appendedSkipped > 0)
    skipped += appendedSkipped;

  var updateDropped = updateQueued - updateWrites.length;
  if (updateDropped > 0) skipped += updateDropped;

  Logger.log('追加対象行数: ' + appendValues.length);
  Logger.log('更新対象行数: ' + updateWrites.length);

  var setValuesRows = appendValues.length + updateWrites.length;
  Logger.log('実際に setValues する行数: ' + setValuesRows);

  for (var u = 0; u < updateWrites.length; u++) {
    var w = updateWrites[u];
    var updateRg = scheduleSheet.getRange(w.rn, 1, 1, numColsPhysical);
    var rowMatrix = [w.line];
    if (rowMatrix.length !== updateRg.getNumRows()) {
      throw new Error(
        'internal: 更新の行数が一致しません (values=' +
          rowMatrix.length +
          ' rangeRows=' +
          updateRg.getNumRows() +
          ')'
      );
    }
    updateRg.setValues(rowMatrix);
  }

  if (appendValues.length > 0) {
    var startRow = scheduleSheet.getLastRow() + 1;
    var appendRg = scheduleSheet.getRange(
      startRow,
      1,
      appendValues.length,
      numColsPhysical
    );

    Logger.log(`appendValues.length=${appendValues.length}`);
    Logger.log(`append startRow=${startRow}`);
    Logger.log(`append rangeRows=${appendRg.getNumRows()}`);
    Logger.log(`numColsPhysical=${numColsPhysical}`);

    if (appendRg.getNumRows() !== appendValues.length) {
      throw new Error(
        'internal: 追記の行数が一致しません (values=' +
          appendValues.length +
          ' rangeRows=' +
          appendRg.getNumRows() +
          ')'
      );
    }

    appendRg.setValues(appendValues);
  }

  var added = appendValues.length;
  var updated = updateWrites.length;

  var needsReview = rows.filter(function (r) {
    return r['マスタ照合状態'] === MASTER_NEEDS_REVIEW;
  }).length;

  Logger.log(
    'UPSERT 完了 added=' +
      added +
      ' updated=' +
      updated +
      ' skipped=' +
      skipped +
      ' needsReview(in rows)=' +
      needsReview +
      ' appendQueued=' +
      appendQueued +
      ' updateQueued=' +
      updateQueued
  );

  return { added: added, updated: updated, skipped: skipped, needsReview: needsReview };
}

/**
 * 1行目のヘッダ配列と mergedObj（列名キー）から、シート列順の1行を作る。
 * ヘッダが空の列は出力を空にする。
 * @param {!Array<*>} headerRowVals シート1行目そのまま
 * @param {!Object<string, *>} mergedObj
 * @return {!Array<?>}
 */
function buildOutputLine_(headerRowVals, mergedObj) {
  var line = [];
  for (var i = 0; i < headerRowVals.length; i++) {
    var hn = normalizeCellValue_(headerRowVals[i]);
    if (!hn) {
      line.push('');
      continue;
    }
    if (Object.prototype.hasOwnProperty.call(mergedObj, hn)) line.push(mergedObj[hn]);
    else line.push('');
  }
  return line;
}

/**
 * UPSERT の複合キーを生成する（予定日キーは `normalizeDateKey_` で yyyy-MM-dd）。
 * @param {!Object<string, *>} row
 * @return {string}
 */
function buildSourceKey_(row) {
  var tz = Session.getScriptTimeZone();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = ss && ss.getSpreadsheetTimeZone();
    if (ssTz) tz = ssTz;
  } catch (ignore) {}

  var sheetNm = normalizeCellValue_(String(row['作成元シート名'] || ''));
  var jd = normalizeCellValue_(String(row['案件名_入力値'] || ''));
  var cell = normalizeCellValue_(String(row['作成元セル'] || ''));
  var dkey = normalizeDateKey_(row['予定日'], tz);

  if (!sheetNm || !jd || !cell || !dkey) return '';
  return sheetNm + '\x1f' + dkey + '\x1f' + jd + '\x1f' + cell;
}

/**
 * 新規運行予定ID（UUID）を発行する。
 * @return {string}
 */
function createScheduleId_() {
  return Utilities.getUuid();
}

/**
 * セル値を文字列へ正規化する。
 * @param {*} value
 * @return {string}
 */
function normalizeCellValue_(value) {
  if (value === null || value === undefined) return '';
  var s = String(value);
  s = s.replace(/^[\u3000\s]+/, '').replace(/[\u3000\s]+$/, '');
  return s.replace(/^[\u0020\t\r\n]+/, '').replace(/[\u0020\t\r\n]+$/, '');
}

/**
 * 予定日列の値を複合キー用 **yyyy-MM-dd** に統一する。
 * `yyyy/MM/dd`・`yyyy-MM-dd` の文字列、スプレッドシート読み込みの `Date`、その他 coerce 可能な入力を想定。
 * @param {*} value
 * @param {string} timeZone IANA TZ
 * @return {string}
 */
function normalizeDateKey_(value, timeZone) {
  try {
    if (value === '' || value === null || value === undefined) return '';
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
    }
    var text = normalizeCellValue_(String(value));
    if (!text) return '';
    var m = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) {
      return (
        m[1] +
        '-' +
        String(Number(m[2])).padStart(2, '0') +
        '-' +
        String(Number(m[3])).padStart(2, '0')
      );
    }
    var coerced = coerceToDate_(value, null, timeZone);
    if (coerced && !isNaN(coerced.getTime())) {
      return Utilities.formatDate(coerced, timeZone, 'yyyy-MM-dd');
    }
    return '';
  } catch (e) {
    return '';
  }
}

/**
 * {@link normalizeDateKey_} へのエイリアス（複合キー・照合用 yyyy-MM-dd）。
 * @param {*} dateValue
 * @param {string} timeZone IANA TZ
 * @return {string}
 */
function formatDateKey_(dateValue, timeZone) {
  return normalizeDateKey_(dateValue, timeZone);
}

/**
 * 行・列は 1起点で指定する。
 * @param {number} row
 * @param {number} col
 * @return {string}
 */
function toA1_(row, col) {
  return columnToLetters_(col) + row;
}

/**
 * @param {number} colNum 列番号 1〜
 * @return {string}
 */
function columnToLetters_(colNum) {
  var letters = '';
  var n = Math.floor(Number(colNum));
  while (n > 0) {
    var m = (n - 1) % 26;
    letters = String.fromCharCode(65 + m) + letters;
    n = Math.floor((n - m - 1) / 26);
  }
  return letters || 'A';
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {!Object<string, number>} ヘッダ名 → 列番号（1起点）
 */
function getSheetHeaderMap_(sheet) {
  var lc = sheet.getLastColumn();
  if (lc < 1) return {};

  var headers = sheet.getRange(1, 1, 1, lc).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = normalizeCellValue_(headers[i]);
    if (!h) continue;
    map[h] = i + 1;
  }
  return map;
}

/** @param {!Object<string, number>} headerMap */
function validateRequiredHeaders_(headerMap) {
  for (var i = 0; i < REQUIRED_SCHEDULE_HEADERS.length; i++) {
    var req = REQUIRED_SCHEDULE_HEADERS[i];
    if (!headerMap[req]) {
      throw new Error(
        '10_配車予定に必須列「' + req + '」がありません（1行目のヘッダーを確認してください）。'
      );
    }
  }
}

/**
 * formatDateKey_ 経由での日付解釈に使用する。
 * @param {*} raw
 * @param {?{year:number, month:number}} yearMonthCtx
 * @param {string} timeZone
 * @return {?Date}
 */
function resolveScheduleDate_(raw, yearMonthCtx, timeZone) {
  if (raw === '' || raw === null || raw === undefined) return null;

  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  }

  if (typeof raw === 'number' && !isNaN(raw)) {
    if (raw > 20000) {
      var epoch = new Date(1899, 11, 30);
      var ms = epoch.getTime() + Math.round(raw * 86400000);
      var d = new Date(ms);
      if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    if (yearMonthCtx && raw >= 1 && raw <= 31 && Math.floor(raw) === raw) {
      return new Date(yearMonthCtx.year, yearMonthCtx.month - 1, raw);
    }
  }

  var s = normalizeCellValue_(raw);
  if (!s) return null;

  var parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    var d2 = new Date(parsed);
    return new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
  }

  var dayNum = parseInt(s, 10);
  if (!isNaN(dayNum) && dayNum >= 1 && dayNum <= 31 && yearMonthCtx) {
    return new Date(yearMonthCtx.year, yearMonthCtx.month - 1, dayNum);
  }

  return null;
}

/**
 * @param {*} value
 * @param {?{year:number, month:number}} ym
 * @param {string} timeZone
 * @return {?Date}
 */
function coerceToDate_(value, ym, timeZone) {
  return resolveScheduleDate_(value, ym, timeZone);
}
