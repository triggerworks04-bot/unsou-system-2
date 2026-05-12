/**
 * 10_配車予定（正本）から 90_配車表表示用（A3印刷・横持ち配車表）を生成する。
 * DB行は増やさず、日跨ぎのみ表示側で複数日に展開する。
 *
 * @fileoverview TARGET_SCHEDULE_SHEET_NAME は dispatch_conversion.gs で定義。
 */

/** @type {string} */
var dispatchView_VIEW_SHEET_NAME = '90_配車表表示用';

/** 設定シート（表示対象年月を参照） */
/** @type {string} */
var dispatchView_SETTINGS_SHEET_NAME = '00_設定';

/** A列で探す項目名（配車表表示年月） */
/** @type {string} */
var dispatchView_SETTINGS_LABEL_DISPLAY_YM = '配車表表示年月';

/** 案件マスタ（90 横列の順と表示名） */
/** @type {string} */
var dispatchView_JOB_MASTER_SHEET_NAME = '08_案件マスター';

/** @type {string} */
var dispatchView_STATUS_CANCELLED = 'キャンセル';

/** 90 の案件列に出さないプレースホルダー（10 上の表記） */
/** @type {string} */
var dispatchView_EXCLUDED_JOB_NAME_FOR_COLUMNS_ = '（案件名なし）';

/** @type {!Array<string>} */
var dispatchView_REQUIRED_SCHEDULE_HEADERS = ['ステータス', '予定日'];

/**
 * @param {string} name
 * @return {boolean}
 */
function dispatchView_isExcludedJobNameForView_(name) {
  var n = normalizeCellValue_(String(name || ''));
  return n === dispatchView_EXCLUDED_JOB_NAME_FOR_COLUMNS_;
}

/** 各行ブロックの左列ラベル（C列データ） */
/** @type {!Array<string>} */
var dispatchView_ROW_LABELS_ = ['担当者', '車両', '積込', '納品', '備考'];

/**
 * メニュー「90_配車表表示用を更新」から実行。
 */
function refreshDispatchViewSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    var scheduleSheet = ss.getSheetByName(TARGET_SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) {
      throw new Error('シート「' + TARGET_SCHEDULE_SHEET_NAME + '」が見つかりません。');
    }
    var viewSheet = ss.getSheetByName(dispatchView_VIEW_SHEET_NAME);
    if (!viewSheet) {
      throw new Error('シート「' + dispatchView_VIEW_SHEET_NAME + '」が見つかりません。');
    }

    var scheduleHeaderMap = getSheetHeaderMap_(scheduleSheet);
    dispatchView_validateScheduleHeaders_(scheduleHeaderMap);

    var tz = dispatchView_getSpreadsheetTimeZone_();

    var schLastRow = scheduleSheet.getLastRow();
    var schLastCol = scheduleSheet.getLastColumn();

    if (schLastRow < 2 || schLastCol < 1) {
      dispatchView_clearAndResetSheet_(viewSheet);
      ui.alert('対象データがありません（' + TARGET_SCHEDULE_SHEET_NAME + '）。表示用シートをクリアしました。');
      return;
    }

    var dataRange = scheduleSheet.getRange(2, 1, schLastRow, schLastCol);
    var rows = dataRange.getValues();

    /** @type {{year:number, month:number}} */
    var targetYm = dispatchView_readTargetYearMonthFromSettings_(ss, tz);

    /** @type {!Array<{dateKey: string, kubun: string, jobName: string, driver: string, vehicle: string, loadLine: string, delivLine: string, remarkLine: string, sortMin: number}>} */
    var atomics = [];

    var ri;
    for (ri = 0; ri < rows.length; ri++) {
      var row = rows[ri];
      if (dispatchView_isCancelledRow_(row, scheduleHeaderMap)) continue;

      var jobForCol = dispatchView_pickFirstNonEmpty2_(
        dispatchView_cell_(row, scheduleHeaderMap, '案件名_入力値'),
        dispatchView_cell_(row, scheduleHeaderMap, '案件名')
      );
      if (!normalizeCellValue_(String(jobForCol || ''))) continue;
      if (dispatchView_isExcludedJobNameForView_(jobForCol)) continue;

      if (!dispatchView_rowMayContributeToTargetMonth_(row, scheduleHeaderMap, tz, targetYm.year, targetYm.month))
        continue;

      var expanded = dispatchView_expandRowToAtomics_(row, scheduleHeaderMap, tz);
      var ei;
      for (ei = 0; ei < expanded.length; ei++) {
        atomics.push(expanded[ei]);
      }
    }

    atomics = dispatchView_filterAtomicsToCalendarMonth_(atomics, targetYm.year, targetYm.month);

    atomics.sort(dispatchView_compareAtomic_);
    atomics = dispatchView_mergeAtomicsSameKey_(atomics);

    var jobCols = dispatchView_buildJobColumnKeysAndLabels_(ss, atomics);

    var titleText = dispatchView_buildTitleFromYearMonth_(targetYm.year, targetYm.month);
    dispatchView_writeGridToSheet_(viewSheet, titleText, jobCols.keys, jobCols.labels, atomics, targetYm);

    ui.alert('更新完了\n' + dispatchView_VIEW_SHEET_NAME + '（横持ち配車表）を出力しました。');
  } catch (e) {
    Logger.log(e.stack || String(e.message));
    ui.alert('エラー: ' + (e.message || String(e)));
  }
}

/**
 * シート全体を結合解除のうえ内容・書式・条件付き書式をクリアし、固定行・列をリセットする。
 * （全体 Range の breakApart() のみでは部分結合が残り merge 競合することがあるため、getMergedRanges() で順に解除する。）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function dispatchView_clearAndResetSheet_(sheet) {
  try {
    var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    /** @type {!Array<GoogleAppsScript.Spreadsheet.Range>} */
    var mergedRanges;
    while ((mergedRanges = fullRange.getMergedRanges()).length > 0) {
      mergedRanges[0].breakApart();
    }
    fullRange.clearContent();
    fullRange.clearFormat();
    sheet.clearConditionalFormatRules();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
  } catch (e) {
    throw new Error('90_配車表表示用の初期化に失敗しました: ' + (e.message || e));
  }
}

/**
 * yyyy-MM-dd 形式で有効か（9999-12-31 を除く）。
 * @param {string} key
 * @return {boolean}
 */
function dispatchView_validYmdKey_(key) {
  var s = normalizeCellValue_(String(key || ''));
  if (!s || s === '9999-12-31') return false;
  return /^(\d{4})-(\d{2})-(\d{2})$/.test(s);
}

/**
 * 00_設定 の「配車表表示年月」と同じ項目名の行の B 列から年月を取得する。
 * シート無し／行無し／B 空／解析失敗 → 実行日時の年月（スプレッドシート TZ）。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} tz
 * @return {{year: number, month: number}}
 */
function dispatchView_readTargetYearMonthFromSettings_(ss, tz) {
  var fallback = dispatchView_currentYearMonthInTz_(tz);
  var sh = ss.getSheetByName(dispatchView_SETTINGS_SHEET_NAME);
  if (!sh) {
    Logger.log('dispatchView: シート「' + dispatchView_SETTINGS_SHEET_NAME + '」がありません。実行日時の年月を使用します。');
    return fallback;
  }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('dispatchView: 00_設定 にデータ行がありません。実行日時の年月を使用します。');
    return fallback;
  }
  var vals = sh.getRange(2, 1, lastRow, 2).getValues();
  var ri;
  for (ri = 0; ri < vals.length; ri++) {
    var label = normalizeCellValue_(String(vals[ri][0] == null ? '' : vals[ri][0]));
    if (label !== dispatchView_SETTINGS_LABEL_DISPLAY_YM) continue;
    var bRaw = vals[ri][1];
    var bStr = normalizeCellValue_(String(bRaw == null ? '' : bRaw));
    if (!bStr) {
      return fallback;
    }
    var parsed = dispatchView_parseDisplayYearMonthValue_(bStr);
    if (!parsed) {
      Logger.log('dispatchView: 表示対象年月の値が不正のため実行日時の年月にフォールバック: ' + bStr);
      return fallback;
    }
    return parsed;
  }
  Logger.log('dispatchView: A列に「' + dispatchView_SETTINGS_LABEL_DISPLAY_YM + '」がありません。実行日時の年月を使用します。');
  return fallback;
}

/**
 * 2026/05・2026-5・2026年5月 等を { year, month } に。
 * @param {string} raw
 * @return {?{year: number, month: number}}
 */
function dispatchView_parseDisplayYearMonthValue_(raw) {
  var s = normalizeCellValue_(String(raw == null ? '' : raw));
  if (!s) return null;
  var m1 = /^(\d{4})年(\d{1,2})月/.exec(s);
  if (m1) {
    return dispatchView_normalizeYearMonthParts_(parseInt(m1[1], 10), parseInt(m1[2], 10));
  }
  var m2 = /^(\d{4})[\/\-](\d{1,2})$/.exec(s);
  if (m2) {
    return dispatchView_normalizeYearMonthParts_(parseInt(m2[1], 10), parseInt(m2[2], 10));
  }
  return null;
}

/**
 * @param {number} y
 * @param {number} mo
 * @return {?{year: number, month: number}}
 */
function dispatchView_normalizeYearMonthParts_(y, mo) {
  if (isNaN(y) || isNaN(mo)) return null;
  if (y < 1900 || y > 2200) return null;
  if (mo < 1 || mo > 12) return null;
  return { year: y, month: mo };
}

/**
 * 対象月に積込日または納品日（単日は予定日）のいずれかが入る行だけ展開する。
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @param {string} tz
 * @param {number} year
 * @param {number} month
 * @return {boolean}
 */
function dispatchView_rowMayContributeToTargetMonth_(row, headerMap, tz, year, month) {
  var planNorm = normalizeDateKey_(dispatchView_getRaw_(row, headerMap, '予定日'), tz);
  if (!dispatchView_validYmdKey_(planNorm)) return false;

  var rawLoadDt = dispatchView_getRaw_(row, headerMap, '積込予定日時');
  var rawDelDt = dispatchView_getRaw_(row, headerMap, '納品予定日時');
  var loadKey = dispatchView_extractDateKeyFromDatetime_(rawLoadDt, tz);
  var delKey = dispatchView_extractDateKeyFromDatetime_(rawDelDt, tz);
  if (!loadKey) loadKey = planNorm;
  if (!delKey) delKey = loadKey || planNorm;
  if (!loadKey || !delKey) return false;

  if (loadKey === delKey) {
    return dispatchView_dateKeyInCalendarMonth_(planNorm, year, month);
  }
  return (
    dispatchView_dateKeyInCalendarMonth_(loadKey, year, month) ||
    dispatchView_dateKeyInCalendarMonth_(delKey, year, month)
  );
}

/**
 * @param {string} tz
 * @return {{year: number, month: number}}
 */
function dispatchView_currentYearMonthInTz_(tz) {
  var txt = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  var m = /^(\d{4})-(\d{2})$/.exec(txt);
  if (m) {
    return { year: parseInt(m[1], 10), month: parseInt(m[2], 10) };
  }
  var d = new Date();
  return { year: d.getFullYear(), month: d.getMonth() + 1 };
}

/**
 * @param {number} year
 * @param {number} month 1〜12
 * @return {number}
 */
function dispatchView_daysInCalendarMonth_(year, month) {
  return new Date(year, month, 0).getDate();
}

/**
 * @param {number} year
 * @param {number} month 1〜12
 * @return {!Array<string>}
 */
function dispatchView_enumerateDateKeysInCalendarMonth_(year, month) {
  var dim = dispatchView_daysInCalendarMonth_(year, month);
  var out = [];
  var d;
  for (d = 1; d <= dim; d++) {
    out.push(year + '-' + String(month).padStart(2, '0') + '-' + String(d).padStart(2, '0'));
  }
  return out;
}

/**
 * @param {string} dk yyyy-MM-dd
 * @param {number} year
 * @param {number} month 1〜12
 * @return {boolean}
 */
function dispatchView_dateKeyInCalendarMonth_(dk, year, month) {
  if (!dispatchView_validYmdKey_(dk)) return false;
  var m = /^(\d{4})-(\d{2})-\d{2}$/.exec(dk);
  if (!m) return false;
  return parseInt(m[1], 10) === year && parseInt(m[2], 10) === month;
}

/**
 * @param {!Array<{dateKey: string}>} atomics
 * @param {number} year
 * @param {number} month
 * @return {!Array<{dateKey: string}>}
 */
function dispatchView_filterAtomicsToCalendarMonth_(atomics, year, month) {
  /** @type {!Array<{dateKey: string}>} */
  var out = [];
  var i;
  for (i = 0; i < atomics.length; i++) {
    var a = atomics[i];
    if (dispatchView_dateKeyInCalendarMonth_(a.dateKey, year, month)) out.push(a);
  }
  return out;
}

/**
 * 対象月に出ている案件キー（10 の案件名_入力値／案件名）を、08_案件マスターの表示順で並べ、見出しは表示名（空なら案件名）。未登録は末尾・名前昇順。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {!Array<{jobName: string}>} atomics
 * @return {{keys: !Array<string>, labels: !Array<string>}}
 */
function dispatchView_buildJobColumnKeysAndLabels_(ss, atomics) {
  /** @type {!Object<boolean>} */
  var need = {};
  var i;
  for (i = 0; i < atomics.length; i++) {
    var k = atomics[i].jobName || '';
    if (k && !dispatchView_isExcludedJobNameForView_(k)) need[k] = true;
  }
  var needKeys = Object.keys(need);
  if (needKeys.length === 0) {
    return { keys: [], labels: [] };
  }

  var masterRows = dispatchView_loadActiveJobMasterOrdered_(ss);
  /** @type {!Array<string>} */
  var outKeys = [];
  /** @type {!Object<boolean>} */
  var fromMaster = {};
  var mi;
  for (mi = 0; mi < masterRows.length; mi++) {
    var mj = masterRows[mi].jobNameKey;
    if (!need[mj]) continue;
    if (fromMaster[mj]) continue;
    fromMaster[mj] = true;
    outKeys.push(mj);
  }

  /** @type {!Array<string>} */
  var unreg = [];
  for (i = 0; i < needKeys.length; i++) {
    if (!fromMaster[needKeys[i]]) unreg.push(needKeys[i]);
  }
  unreg.sort();
  for (i = 0; i < unreg.length; i++) outKeys.push(unreg[i]);

  /** @type {!Object<string>} */
  var headerByKey = {};
  for (mi = 0; mi < masterRows.length; mi++) {
    var mk = masterRows[mi].jobNameKey;
    if (!Object.prototype.hasOwnProperty.call(headerByKey, mk)) headerByKey[mk] = masterRows[mi].headerLabel;
  }

  /** @type {!Array<string>} */
  var labels = [];
  for (i = 0; i < outKeys.length; i++) {
    var kk = outKeys[i];
    if (Object.prototype.hasOwnProperty.call(headerByKey, kk) && headerByKey[kk]) {
      labels.push(headerByKey[kk]);
    } else {
      labels.push(kk);
    }
  }

  return { keys: outKeys, labels: labels };
}

/**
 * 有効な行を表示順で返す。シートや「案件名」列が無いときは []。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {!Array<{jobNameKey: string, headerLabel: string, sortOrder: number, rowIx: number}>}
 */
function dispatchView_loadActiveJobMasterOrdered_(ss) {
  var sh = ss.getSheetByName(dispatchView_JOB_MASTER_SHEET_NAME);
  if (!sh) {
    Logger.log(
      'dispatchView: シート「' + dispatchView_JOB_MASTER_SHEET_NAME + '」がないため、マスタ順は使いません。'
    );
    return [];
  }
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  var headerMap;
  try {
    headerMap = getSheetHeaderMap_(sh);
  } catch (ignored) {
    return [];
  }
  if (!headerMap['案件名']) {
    Logger.log('dispatchView: 08_案件マスターに「案件名」列がありません。');
    return [];
  }

  var vals = sh.getRange(2, 1, lr, lc).getValues();
  /** @type {!Array<{jobNameKey: string, headerLabel: string, sortOrder: number, rowIx: number}>} */
  var buf = [];
  var ri;
  for (ri = 0; ri < vals.length; ri++) {
    var row = vals[ri];
    if (!dispatchView_masterRowIsActive_(row, headerMap)) continue;
    var jn = dispatchView_cell_(row, headerMap, '案件名');
    if (!normalizeCellValue_(String(jn || ''))) continue;
    var dispNm = dispatchView_cell_(row, headerMap, '表示名');
    var hdr = normalizeCellValue_(String(dispNm || '')) ? dispNm : jn;
    var ordRaw = headerMap['表示順'] ? row[headerMap['表示順'] - 1] : '';
    var sortOrder = dispatchView_parseMasterDisplayOrder_(ordRaw);
    buf.push({ jobNameKey: jn, headerLabel: hdr, sortOrder: sortOrder, rowIx: ri });
  }
  buf.sort(dispatchView_compareMasterRows_);
  return buf;
}

/**
 * @param {!Object<string, *>} a
 * @param {!Object<string, *>} b
 * @return {number}
 */
function dispatchView_compareMasterRows_(a, b) {
  if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
  return a.rowIx - b.rowIx;
}

/**
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @return {boolean}
 */
function dispatchView_masterRowIsActive_(row, headerMap) {
  if (!headerMap['有効フラグ']) return true;
  return dispatchView_truthyEffectiveFlag_(dispatchView_getRaw_(row, headerMap, '有効フラグ'));
}

/**
 * @param {*} v
 * @return {boolean}
 */
function dispatchView_truthyEffectiveFlag_(v) {
  if (v === true) return true;
  if (v === false) return false;
  if (typeof v === 'number' && !isNaN(v)) return v !== 0;
  var s = normalizeCellValue_(String(v == null ? '' : v)).toLowerCase();
  if (!s) return false;
  return (
    s === 'true' ||
    s === '1' ||
    s === 'yes' ||
    s === 'はい' ||
    s === '○' ||
    s === '済'
  );
}

/**
 * @param {*} v
 * @return {number}
 */
function dispatchView_parseMasterDisplayOrder_(v) {
  if (v === '' || v === null || v === undefined) return 999999;
  var n = Number(v);
  if (!isNaN(n) && isFinite(n)) return n;
  return 999999;
}

/**
 * @param {number} year
 * @param {number} month 1〜12
 */
function dispatchView_buildTitleFromYearMonth_(year, month) {
  var moNum = Number(month);
  var yNum = Number(year);
  if (!yNum || moNum < 1 || moNum > 12) return '配車表';
  return yNum + '年' + moNum + '月 配車表';
}

/**
 * yyyy-MM-dd をグレゴリオ暦の暦日として解釈し、曜日 0=日..6=土（Session の TZ に依存しない）。
 * @param {string} ymdKey
 * @return {number} 無効時は -1
 */
function dispatchView_getIsoYmdWeekdayIndex_(ymdKey) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizeCellValue_(String(ymdKey || '')));
  if (!m) return -1;
  var y = parseInt(m[1], 10);
  var mo = parseInt(m[2], 10);
  var d = parseInt(m[3], 10);
  var t = Date.UTC(y, mo - 1, d, 12, 0, 0);
  var dt = new Date(t);
  if (isNaN(dt.getTime())) return -1;
  return dt.getUTCDay();
}

/**
 * @param {string} ymdKey
 * @return {boolean}
 */
function dispatchView_isSundayDateKey_(ymdKey) {
  return dispatchView_getIsoYmdWeekdayIndex_(ymdKey) === 0;
}

/**
 * @param {!Object<string, number>} headerMap
 */
function dispatchView_validateScheduleHeaders_(headerMap) {
  var i;
  for (i = 0; i < dispatchView_REQUIRED_SCHEDULE_HEADERS.length; i++) {
    var n = dispatchView_REQUIRED_SCHEDULE_HEADERS[i];
    if (!headerMap[n]) {
      throw new Error(
        TARGET_SCHEDULE_SHEET_NAME +
          ' に必須列「' +
          n +
          '」がありません（1行目のヘッダーを確認してください）。'
      );
    }
  }
}

/**
 * @return {string}
 */
function dispatchView_getSpreadsheetTimeZone_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = ss && ss.getSpreadsheetTimeZone();
    if (ssTz) return ssTz;
  } catch (ignore) {}
  return Session.getScriptTimeZone() || 'Asia/Tokyo';
}

/**
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @return {boolean}
 */
function dispatchView_isCancelledRow_(row, headerMap) {
  var st = dispatchView_cell_(row, headerMap, 'ステータス');
  return st === dispatchView_STATUS_CANCELLED;
}

/**
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @param {string} name
 * @return {string}
 */
function dispatchView_cell_(row, headerMap, name) {
  var col = headerMap[name];
  if (!col) return '';
  return normalizeCellValue_(row[col - 1]);
}

/**
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @param {string} name
 * @return {*}
 */
function dispatchView_getRaw_(row, headerMap, name) {
  var col = headerMap[name];
  if (!col) return '';
  return row[col - 1];
}

/**
 * 備考欄（/ 区切り）：荷種・重量・運賃・高速代・備考。optional で語を追加してから結合。
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @param {string} extraRemarkToken 追加のみ（単一日跨ぎ用）
 * @return {string}
 */
function dispatchView_buildSlashRemarkParts_(row, headerMap, extraRemarkToken) {
  var cols = ['荷種', '重量', '運賃', '高速代', '備考'];
  var parts = [];
  var ci;
  for (ci = 0; ci < cols.length; ci++) {
    var nm = cols[ci];
    var v =
      nm === '備考'
        ? dispatchView_pickFirstNonEmpty2_(
            dispatchView_cell_(row, headerMap, '備考'),
            dispatchView_cell_(row, headerMap, '摘要・ルート等')
          )
        : dispatchView_cell_(row, headerMap, nm);
    v = normalizeCellValue_(String(v || ''));
    if (v) parts.push(v);
  }
  var ex = normalizeCellValue_(String(extraRemarkToken || ''));
  if (ex) {
    if (parts.join('').indexOf(ex) === -1) parts.push(ex);
  }
  return parts.join(' / ');
}

/**
 * HH:mm と地を結合（空要素は省略）。
 * @param {string} prefix 例 翌日 / 前日 / N日後
 * @param {string} hm
 * @param {string} place
 * @return {string}
 */
function dispatchView_formatLoadDelivCell_(prefix, hm, place) {
  var p = normalizeCellValue_(String(prefix || ''));
  var h = normalizeCellValue_(String(hm || ''));
  var pl = normalizeCellValue_(String(place || ''));
  var bits = [];
  if (p) bits.push(p);
  if (h) bits.push(h);
  if (pl) bits.push(pl);
  return bits.join(' ');
}

/**
 * @param {!Array<*>} row
 * @param {!Object<string, number>} headerMap
 * @param {string} tz
 * @return {!Array<!{dateKey: string, kubun: string, jobName: string, driver: string, vehicle: string, loadLine: string, delivLine: string, remarkLine: string, sortMin: number}>}
 */
function dispatchView_expandRowToAtomics_(row, headerMap, tz) {
  var jobName = dispatchView_pickFirstNonEmpty2_(
    dispatchView_cell_(row, headerMap, '案件名_入力値'),
    dispatchView_cell_(row, headerMap, '案件名')
  );
  if (!normalizeCellValue_(String(jobName || ''))) {
    return [];
  }

  var planNorm = normalizeDateKey_(dispatchView_getRaw_(row, headerMap, '予定日'), tz);
  if (!dispatchView_validYmdKey_(planNorm)) {
    return [];
  }

  var kubun = dispatchView_cell_(row, headerMap, '案件種別');
  var driver = dispatchView_pickFirstNonEmpty2_(
    dispatchView_cell_(row, headerMap, '運転者名_入力値'),
    dispatchView_cell_(row, headerMap, '運転者ID')
  );
  var vehicle = dispatchView_pickFirstNonEmpty2_(
    dispatchView_cell_(row, headerMap, '車両呼称_入力値'),
    dispatchView_cell_(row, headerMap, '車両ID')
  );

  var loadHm = dispatchView_extractTimePrimaryOrFallback_(
    dispatchView_getRaw_(row, headerMap, '積込予定日時'),
    dispatchView_getRaw_(row, headerMap, '出発予定時刻'),
    tz
  );
  var deliHm = dispatchView_extractTimePrimaryOrFallback_(
    dispatchView_getRaw_(row, headerMap, '納品予定日時'),
    dispatchView_getRaw_(row, headerMap, '到着予定時刻'),
    tz
  );
  var loadPlace = dispatchView_pickFirstNonEmpty2_(
    dispatchView_cell_(row, headerMap, '積込地'),
    dispatchView_cell_(row, headerMap, '出発地')
  );
  var deliPlace = dispatchView_pickFirstNonEmpty2_(
    dispatchView_cell_(row, headerMap, '納品地'),
    dispatchView_cell_(row, headerMap, '目的地')
  );

  var rawLoadDt = dispatchView_getRaw_(row, headerMap, '積込予定日時');
  var rawDelDt = dispatchView_getRaw_(row, headerMap, '納品予定日時');
  var loadKey = dispatchView_extractDateKeyFromDatetime_(rawLoadDt, tz);
  var delKey = dispatchView_extractDateKeyFromDatetime_(rawDelDt, tz);
  if (!loadKey) loadKey = planNorm;
  if (!delKey) delKey = loadKey || planNorm;

  var loadLinePlain = dispatchView_joinHmPlace_(loadHm, loadPlace);
  var deliLinePlain = dispatchView_joinHmPlace_(deliHm, deliPlace);

  if (!loadKey || !delKey || loadKey === delKey) {
    return [
      dispatchView_stripDisplayFromAtomic_({
        dateKey: planNorm,
        kubun: kubun,
        jobName: jobName,
        driver: driver,
        vehicle: vehicle,
        loadLine: loadLinePlain,
        delivLine: deliLinePlain,
        remarkLine: dispatchView_buildSlashRemarkParts_(row, headerMap, ''),
        sortMin: dispatchView_hmToMinutes_(loadHm),
      }),
    ];
  }

  var diffDays = dispatchView_calendarDaysBetween_(loadKey, delKey);
  if (diffDays <= 0) {
    return [
      dispatchView_stripDisplayFromAtomic_({
        dateKey: planNorm,
        kubun: kubun,
        jobName: jobName,
        driver: driver,
        vehicle: vehicle,
        loadLine: loadLinePlain,
        delivLine: deliLinePlain,
        remarkLine: dispatchView_buildSlashRemarkParts_(row, headerMap, ''),
        sortMin: dispatchView_hmToMinutes_(loadHm),
      }),
    ];
  }

  var relF = dispatchView_relativeDayLabelForward_(diffDays);
  var relB = dispatchView_relativeDayLabelBackward_(diffDays);

  var loadRemark = dispatchView_buildSlashRemarkParts_(row, headerMap, '日跨ぎ');
  var delRemark = dispatchView_buildSlashRemarkParts_(row, headerMap, '前日から継続');

  var deliForLoadSide = dispatchView_formatLoadDelivCell_(relF, deliHm, deliPlace);
  var loadForDelSide = dispatchView_formatLoadDelivCell_(relB, loadHm, loadPlace);

  var loadAtomic = {
    dateKey: loadKey,
    kubun: kubun,
    jobName: jobName,
    driver: driver,
    vehicle: vehicle,
    loadLine: loadLinePlain,
    delivLine: deliForLoadSide,
    remarkLine: loadRemark,
    sortMin: dispatchView_hmToMinutes_(loadHm),
  };
  var delAtomic = {
    dateKey: delKey,
    kubun: kubun,
    jobName: jobName,
    driver: driver,
    vehicle: vehicle,
    loadLine: loadForDelSide,
    delivLine: deliLinePlain,
    remarkLine: delRemark,
    sortMin: dispatchView_hmToMinutes_(deliHm),
  };
  return [dispatchView_stripDisplayFromAtomic_(loadAtomic), dispatchView_stripDisplayFromAtomic_(delAtomic)];
}

/**
 * @param {!Object} a
 * @return {!Object}
 */
function dispatchView_stripDisplayFromAtomic_(a) {
  var o = {
    dateKey: a.dateKey,
    kubun: a.kubun,
    jobName: a.jobName,
    driver: a.driver,
    vehicle: a.vehicle,
    loadLine: a.loadLine,
    delivLine: a.delivLine,
    remarkLine: a.remarkLine,
    sortMin: a.sortMin,
  };
  return o;
}

/**
 * @param {string} hm
 * @param {string} place
 * @return {string}
 */
function dispatchView_joinHmPlace_(hm, place) {
  var h = normalizeCellValue_(String(hm || ''));
  var p = normalizeCellValue_(String(place || ''));
  if (h && p) return h + ' ' + p;
  return h || p;
}

/**
 * @param {!Array<{dateKey: string, kubun: string, jobName: string, driver: string, vehicle: string, loadLine: string, delivLine: string, remarkLine: string, sortMin: number}>} atomicsSorted
 */
function dispatchView_compareAtomic_(a, b) {
  if (a.dateKey !== b.dateKey) return a.dateKey < b.dateKey ? -1 : 1;
  var ka = a.kubun || '';
  var kb = b.kubun || '';
  if (ka !== kb) return ka < kb ? -1 : ka > kb ? 1 : 0;
  if (a.sortMin !== b.sortMin) return a.sortMin - b.sortMin;
  var ja = a.jobName || '';
  var jb = b.jobName || '';
  return ja < jb ? -1 : ja > jb ? 1 : 0;
}

/**
 * 同一 表示日・案件名・案件種別 の原子を縦につなぐ。
 */
function dispatchView_mergeAtomicsSameKey_(atomicsSorted) {
  if (atomicsSorted.length === 0) return atomicsSorted;
  var out = [];
  var cur = atomicsSorted[0];
  var buf = dispatchView_cloneAtomicFields_(cur);
  var nk;
  var i;
  var n;
  for (i = 1; i < atomicsSorted.length; i++) {
    n = atomicsSorted[i];
    nk = n.dateKey + '\x1f' + (n.jobName || '') + '\x1f' + (n.kubun || '');
    var ck =
      buf.dateKey + '\x1f' + (buf.jobName || '') + '\x1f' + (buf.kubun || '');
    if (nk !== ck) {
      out.push(buf);
      buf = dispatchView_cloneAtomicFields_(n);
    } else {
      buf.driver = dispatchView_appendNewline_(buf.driver, n.driver);
      buf.vehicle = dispatchView_appendNewline_(buf.vehicle, n.vehicle);
      buf.loadLine = dispatchView_appendNewline_(buf.loadLine, n.loadLine);
      buf.delivLine = dispatchView_appendNewline_(buf.delivLine, n.delivLine);
      buf.remarkLine = dispatchView_appendNewline_(buf.remarkLine, n.remarkLine);
      if (n.sortMin < buf.sortMin) buf.sortMin = n.sortMin;
    }
  }
  out.push(buf);
  return out;
}

/**
 * @param {!Object<string, *>} a
 */
function dispatchView_cloneAtomicFields_(a) {
  return {
    dateKey: a.dateKey,
    kubun: a.kubun,
    jobName: a.jobName,
    driver: a.driver,
    vehicle: a.vehicle,
    loadLine: a.loadLine,
    delivLine: a.delivLine,
    remarkLine: a.remarkLine,
    sortMin: a.sortMin,
  };
}

/**
 * @param {string} a
 * @param {string} b
 * @return {string}
 */
function dispatchView_appendNewline_(a, b) {
  var x = normalizeCellValue_(String(a || ''));
  var y = normalizeCellValue_(String(b || ''));
  if (!y) return x;
  if (!x) return y;
  return x + '\n' + y;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} viewSheet
 * @param {string} titleText
 * @param {!Array<string>} jobColumnKeys 照合キー（10 の案件名一致用・atomic.jobName と同一文字列）
 * @param {!Array<string>} jobColumnLabels 2行目見出し（表示名）
 * @param {!Array<{dateKey: string, kubun: string, jobName: string, driver: string, vehicle: string, loadLine: string, delivLine: string, remarkLine: string, sortMin: number}>} atomics
 * @param {{year: number, month: number}} targetYm 対象月（1〜12）
 */
function dispatchView_writeGridToSheet_(viewSheet, titleText, jobColumnKeys, jobColumnLabels, atomics, targetYm) {
  dispatchView_clearAndResetSheet_(viewSheet);

  if (jobColumnKeys.length === 0) {
    var fb = dispatchView_inferJobColumnKeysAndLabelsFromAtomics_(atomics);
    jobColumnKeys = fb.keys;
    jobColumnLabels = fb.labels;
  }
  while (jobColumnLabels.length < jobColumnKeys.length) jobColumnLabels.push('');
  jobColumnLabels = jobColumnLabels.slice(0, jobColumnKeys.length);

  var colCount = Math.max(3 + jobColumnKeys.length, 3);

  var dateBuckets = dispatchView_bucketAtomicsByDate_(atomics);
  var sortedDates = dispatchView_enumerateDateKeysInCalendarMonth_(targetYm.year, targetYm.month);

  /** @type {!Array<!Array<*>>} */
  var grid = [];
  var titleRow = [];
  var ji;
  for (ji = 0; ji < colCount; ji++) titleRow.push(ji === 0 ? titleText : '');
  grid.push(titleRow);

  var headRow = ['日付', '曜日', '項目'];
  for (ji = 0; ji < jobColumnKeys.length; ji++) headRow.push(jobColumnLabels[ji] || jobColumnKeys[ji]);
  while (headRow.length < colCount) headRow.push('');
  grid.push(headRow);

  var di;
  for (di = 0; di < sortedDates.length; di++) {
    var dk = sortedDates[di];
    var list = dateBuckets[dk] || [];
    list.sort(dispatchView_compareAtomic_);
    var disp = dispatchView_formatDisplayDateFromKey_(dk);
    var wd = dispatchView_weekdayKanjiOneCharFromDateKey_(dk);
    var li;
    for (li = 0; li < dispatchView_ROW_LABELS_.length; li++) {
      var line = [];
      if (li === 0) line.push(disp);
      else line.push('');
      if (li === 0) line.push(wd);
      else line.push('');
      line.push(dispatchView_ROW_LABELS_[li]);
      var hj;
      for (hj = 0; hj < jobColumnKeys.length; hj++) {
        var jn = jobColumnKeys[hj];
        var cell = dispatchView_cellForJobAndLabel_(list, jn, li);
        line.push(cell);
      }
      while (line.length < colCount) line.push('');
      grid.push(line);
    }
  }

  var numRows = grid.length;
  var range = viewSheet.getRange(1, 1, numRows, colCount);
  range.setValues(grid);

  var titleRowRange = viewSheet.getRange(1, 1, 1, colCount);
  titleRowRange.breakApart();
  titleRowRange
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#2f3e5e')
    .setFontColor('#ffffff')
    .setVerticalAlignment('middle')
    .setWrap(false);
  viewSheet.getRange(1, 1).setHorizontalAlignment('left');

  /** ヘッダー行（日付・曜日・項目）：タイトルと区別する淡色 */
  var headerRowFill = '#d9e1f2';
  viewSheet
    .getRange(2, 1, 2, colCount)
    .setBackground(headerRowFill)
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);
  viewSheet.getRange(2, 1, 2, 3).setHorizontalAlignment('center');

  /** 本文は範囲全体を白にしたうえで、日曜の5行ブロックのみ薄ピンク（TZずれで平日が日曜色になるのを防ぐため暦日はUTCで判定） */
  var sundayFill = '#fdecec';

  if (numRows > 2) {
    var body = viewSheet.getRange(3, 1, numRows, colCount);
    body.setBackground('#ffffff');
    body.setFontSize(9);
    body.setWrap(true);
    body.setVerticalAlignment('middle');
    body.setBorder(true, true, true, true, true, true, '#333333', SpreadsheetApp.BorderStyle.SOLID);
  }
  viewSheet.getRange(1, 1, 2, colCount).setBorder(true, true, true, true, null, null, '#333333', SpreadsheetApp.BorderStyle.SOLID);

  var br;
  for (br = 0; br < sortedDates.length; br++) {
    var dk2 = sortedDates[br];
    var rTop = 3 + br * 5;
    if (rTop <= numRows) {
      viewSheet.getRange(rTop, 1, rTop, 2).setHorizontalAlignment('center');
    }
    if (dispatchView_isSundayDateKey_(dk2) && rTop + 4 <= numRows) {
      viewSheet.getRange(rTop, 1, rTop + 4, colCount).setBackground(sundayFill);
    }
  }

  /** 前回より行数が減った場合など、グリッド下に残る書式のピンクを消す（まず白で上書き） */
  var lastOnSheet = viewSheet.getLastRow();
  if (lastOnSheet > numRows) {
    viewSheet.getRange(numRows + 1, 1, lastOnSheet, colCount).setBackground('#ffffff');
  }

  viewSheet.setColumnWidth(1, 100);
  viewSheet.setColumnWidth(2, 56);
  viewSheet.setColumnWidth(3, 56);
  for (ji = 0; ji < jobColumnKeys.length; ji++) {
    viewSheet.setColumnWidth(4 + ji, 140);
  }

  viewSheet.setFrozenRows(2);
  viewSheet.setFrozenColumns(3);

  if (numRows >= 3) viewSheet.setRowHeights(3, numRows - 2, 26);

  try {
    viewSheet
      .getPageSetup()
      .setPaperSize(SpreadsheetApp.PaperSize.A3)
      .setOrientation(SpreadsheetApp.PrintOrientation.LANDSCAPE);
  } catch (ignore) {}
}

/**
 * マスタ無し時のフォールバック：キー＝見出し（名前昇順・重複除去）。
 * @param {!Array<{jobName: string}>} atomics
 * @return {{keys: !Array<string>, labels: !Array<string>}}
 */
function dispatchView_inferJobColumnKeysAndLabelsFromAtomics_(atomics) {
  /** @type {!Object<boolean>} */
  var seen = {};
  /** @type {!Array<string>} */
  var order = [];
  var i;
  for (i = 0; i < atomics.length; i++) {
    var jn = atomics[i].jobName || '';
    if (!jn || seen[jn] || dispatchView_isExcludedJobNameForView_(jn)) continue;
    seen[jn] = true;
    order.push(jn);
  }
  order.sort();
  return { keys: order, labels: order.slice() };
}

/**
 * @param {!Array<{dateKey: string}>} atomics
 */
function dispatchView_bucketAtomicsByDate_(atomics) {
  /** @type {!Object<string, !Array<*>>} */
  var b = {};
  var i;
  for (i = 0; i < atomics.length; i++) {
    var a = atomics[i];
    var k = a.dateKey;
    if (!b[k]) b[k] = [];
    b[k].push(a);
  }
  return b;
}

/**
 * @param {!Array<{jobName: string, driver: string, vehicle: string, loadLine: string, delivLine: string, remarkLine: string}>} list
 * @param {string} jobName
 * @param {number} labelIndex 0..4
 * @return {string}
 */
function dispatchView_cellForJobAndLabel_(list, jobName, labelIndex) {
  var parts = [];
  var i;
  for (i = 0; i < list.length; i++) {
    if ((list[i].jobName || '') !== jobName) continue;
    var v = '';
    if (labelIndex === 0) v = list[i].driver || '';
    else if (labelIndex === 1) v = list[i].vehicle || '';
    else if (labelIndex === 2) v = list[i].loadLine || '';
    else if (labelIndex === 3) v = list[i].delivLine || '';
    else v = list[i].remarkLine || '';
    v = normalizeCellValue_(String(v));
    if (v) parts.push(v);
  }
  return parts.join('\n');
}

/**
 * @param {*} a
 * @param {*} b
 * @return {string}
 */
function dispatchView_pickFirstNonEmpty2_(a, b) {
  var s1 = normalizeCellValue_(a == null ? '' : String(a));
  if (s1) return s1;
  return normalizeCellValue_(b == null ? '' : String(b));
}

/**
 * @param {string} ymdKey yyyy-MM-dd
 * @return {string} yyyy/MM/dd
 */
function dispatchView_formatDisplayDateFromKey_(ymdKey) {
  var s = normalizeCellValue_(ymdKey == null ? '' : String(ymdKey));
  if (!s) return '';
  return s.replace(/-/g, '/');
}

/**
 * yyyy-MM-dd から曜日を1文字（月火水木金土日）で返す。
 * @param {string} ymdKey
 * @return {string}
 */
function dispatchView_weekdayKanjiOneCharFromDateKey_(ymdKey) {
  var wd = dispatchView_getIsoYmdWeekdayIndex_(ymdKey);
  if (wd < 0) return '';
  var labels = ['日', '月', '火', '水', '木', '金', '土'];
  return labels[wd];
}

/**
 * @param {*} value
 * @param {string} tz
 * @return {string} yyyy-MM-dd or ''
 */
function dispatchView_extractDateKeyFromDatetime_(value, tz) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return normalizeDateKey_(value, tz);
  }
  var text = normalizeCellValue_(String(value));
  if (!text) return '';
  var combined = text.match(
    /^(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/
  );
  if (combined) {
    return (
      combined[1] +
      '-' +
      String(Number(combined[2])).padStart(2, '0') +
      '-' +
      String(Number(combined[3])).padStart(2, '0')
    );
  }
  return normalizeDateKey_(value, tz);
}

/**
 * @param {string} fromKey yyyy-MM-dd
 * @param {string} toKey yyyy-MM-dd
 * @return {number}
 */
function dispatchView_calendarDaysBetween_(fromKey, toKey) {
  var a = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizeCellValue_(String(fromKey || '')));
  var b = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizeCellValue_(String(toKey || '')));
  if (!a || !b) return 0;
  var t0 = new Date(parseInt(a[1], 10), parseInt(a[2], 10) - 1, parseInt(a[3], 10), 12, 0, 0).getTime();
  var t1 = new Date(parseInt(b[1], 10), parseInt(b[2], 10) - 1, parseInt(b[3], 10), 12, 0, 0).getTime();
  return Math.round((t1 - t0) / 86400000);
}

/**
 * @param {number} n
 * @return {string}
 */
function dispatchView_relativeDayLabelForward_(n) {
  if (n <= 0) return '';
  if (n === 1) return '翌日';
  return n + '日後';
}

/**
 * @param {number} n
 * @return {string}
 */
function dispatchView_relativeDayLabelBackward_(n) {
  if (n <= 0) return '';
  if (n === 1) return '前日';
  return n + '日前';
}

/**
 * @param {*} raw
 * @param {string} tz
 * @return {string}
 */
function dispatchView_sortDateKey_(raw, tz) {
  var k = normalizeDateKey_(raw, tz);
  return k || '9999-12-31';
}

/**
 * @param {string} hm
 * @return {number}
 */
function dispatchView_hmToMinutes_(hm) {
  var s = normalizeCellValue_(hm == null ? '' : String(hm));
  if (!s) return 99999;
  var m = /^(\d{1,2}):(\d{2})$/.exec(s);
  if (!m) return 99999;
  var h = parseInt(m[1], 10);
  var mi = parseInt(m[2], 10);
  if (h < 0 || h > 23 || mi < 0 || mi > 59) return 99999;
  return h * 60 + mi;
}

/**
 * @param {number} n
 * @return {string}
 */
function dispatchView_pad2_(n) {
  var x = Math.floor(Number(n));
  var t = String(x);
  return t.length >= 2 ? t : ('0' + t).slice(-2);
}

/**
 * @param {*} primary
 * @param {*} fallback
 * @param {string} tz
 * @return {string}
 */
function dispatchView_extractTimePrimaryOrFallback_(primary, fallback, tz) {
  var hm = dispatchView_extractHmFromDatetime_(primary, tz);
  if (hm) return hm;
  return dispatchView_extractHmFromTimeLike_(fallback, tz);
}

/**
 * @param {*} value
 * @param {string} tz
 * @return {string}
 */
function dispatchView_extractHmFromDatetime_(value, tz) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'HH:mm');
  }
  var text = normalizeCellValue_(String(value));
  if (!text) return '';
  var combined = text.match(
    /^(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/
  );
  if (combined) {
    return (
      dispatchView_pad2_(parseInt(combined[4], 10)) + ':' + dispatchView_pad2_(parseInt(combined[5], 10))
    );
  }
  var timeOnly = text.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (timeOnly) {
    return dispatchView_pad2_(parseInt(timeOnly[1], 10)) + ':' + dispatchView_pad2_(parseInt(timeOnly[2], 10));
  }
  return '';
}

/**
 * @param {*} value
 * @param {string} tz
 * @return {string}
 */
function dispatchView_extractHmFromTimeLike_(value, tz) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    var y = value.getFullYear();
    var mo = value.getMonth();
    var da = value.getDate();
    if (y === 1899 && mo === 11 && da === 30) {
      var h0 = value.getHours();
      var mi0 = value.getMinutes();
      if (h0 === 0 && mi0 === 0 && value.getSeconds() === 0) return '';
    }
    return Utilities.formatDate(value, tz, 'HH:mm');
  }
  if (typeof value === 'number' && !isNaN(value)) {
    if (value > 0 && value < 1) {
      var ms = Math.round(value * 86400000);
      var h = Math.floor(ms / 3600000) % 24;
      var mi = Math.floor((ms % 3600000) / 60000);
      return dispatchView_pad2_(h) + ':' + dispatchView_pad2_(mi);
    }
  }
  var text = normalizeCellValue_(String(value));
  if (!text) return '';
  var m = text.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (m) {
    return dispatchView_pad2_(parseInt(m[1], 10)) + ':' + dispatchView_pad2_(parseInt(m[2], 10));
  }
  return '';
}
