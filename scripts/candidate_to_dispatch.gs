/**
 * 11_案件候補（決定・確認済・未反映）を 10_配車予定へ追記し、11 側を更新する。
 * 10_配車予定.予定日は積込日優先（無ければ納品日）。積込/納品予定日時は表示値ベースで **yyyy/MM/dd HH:mm** 文字列のみ（時刻空なら空欄）。
 * 候補反映では出発・到着「予定時刻」列には書込まず、拡張列の積込/納品予定日時に集約する。
 *
 * 既存の月別横持ち→10_配車予定変換（dispatch_conversion.gs）とは独立。
 * 年なしの積込/納品日（5/11 等）は 00_設定.配車表表示年月の年で補完（dispatch_view の読み取りと同一）。
 * ヘルパーは candidateToDispatch_ 接頭辞で既存と衝突しない。
 */

/** @type {string} */
var candidateToDispatch_CANDIDATE_SHEET_NAME = '11_案件候補';

/** @type {string} */
var candidateToDispatch_STATUS_DECIDED = '決定';

/** @type {string} */
var candidateToDispatch_CONFIRM_DONE = '確認済';

/** @type {string} */
var candidateToDispatch_STATUS_DISPATCHED = '配車済';

/** @type {string} 10_配車予定 拡張・候補反映時の配車ステータス */
var candidateToDispatch_SCHEDULE_ASSIGN_STATUS = '未配車';

/** @type {!Array<string>} 11 側で必須とするヘッダー（1行目） */
var candidateToDispatch_REQUIRED_CANDIDATE_HEADERS = [
  '案件ステータス',
  '確認ステータス',
  '配車予定ID',
  '案件名',
  'エラー内容',
  '更新日時',
];

/**
 * メニューから実行。決定・確認済・配車予定ID未設定の行のみ 10 へ反映する。
 */
function reflectDecidedCandidatesToSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    var scheduleSheet = ss.getSheetByName(TARGET_SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) {
      throw new Error('シート「' + TARGET_SCHEDULE_SHEET_NAME + '」が見つかりません。');
    }
    var scheduleHeaderMap = getSheetHeaderMap_(scheduleSheet);
    validateRequiredHeaders_(scheduleHeaderMap);

    var candidateSheet = ss.getSheetByName(candidateToDispatch_CANDIDATE_SHEET_NAME);
    if (!candidateSheet) {
      throw new Error('シート「' + candidateToDispatch_CANDIDATE_SHEET_NAME + '」が見つかりません。');
    }

    var dataRange = candidateSheet.getDataRange();
    var candData = dataRange.getValues();
    var candDisplay = dataRange.getDisplayValues();
    if (!candData || candData.length < 2) {
      ui.alert('対象データがありません（' + candidateToDispatch_CANDIDATE_SHEET_NAME + '）。');
      return;
    }

    var candHeaderRow = candData[0];
    candidateToDispatch_validateCandidateHeaders_(candHeaderRow);

    var candCol = candidateToDispatch_headerNameToIndex_(candHeaderRow);
    var lastColSch = scheduleSheet.getLastColumn();
    var headerRowVals = scheduleSheet.getRange(1, 1, 1, lastColSch).getValues()[0];

    var added = 0;
    var skipped = 0;
    var errors = 0;

    var tz0 = candidateToDispatch_getSpreadsheetTimeZone_();
    /** 年なしの月日（5/11 等）を解釈するときの補完年（00_設定.配車表表示年月 → 実行日時の年月） */
    var fallbackYear = candidateToDispatch_getFallbackYear_(ss, tz0);

    for (var r = 1; r < candData.length; r++) {
      var row = candData[r];
      var sheetRowNum = r + 1;

      if (!candidateToDispatch_rowMatchesTransferFilter_(row, candCol)) {
        skipped++;
        continue;
      }

      try {
        var rowDisplay = r < candDisplay.length ? candDisplay[r] : row;
        var timestampText = candidateToDispatch_nowText_();
        var merged = candidateToDispatch_buildMergedFromCandidateRow_(
          row,
          rowDisplay,
          candCol,
          sheetRowNum,
          timestampText,
          fallbackYear
        );
        candidateToDispatch_appendScheduleRow_(scheduleSheet, headerRowVals, merged);

        var newId = merged['運行予定ID'];
        candidateToDispatch_patchCandidateAfterSuccess_(
          candidateSheet,
          sheetRowNum,
          candCol,
          newId,
          timestampText
        );
        added++;
      } catch (e) {
        errors++;
        var msg = e && e.message ? String(e.message) : String(e);
        candidateToDispatch_writeCandidateError_(candidateSheet, sheetRowNum, candCol, msg);
        Logger.log('candidate row ' + sheetRowNum + ': ' + (e.stack || msg));
      }
    }

    var summary =
      '反映完了\n' +
      '追加（10_配車予定）: ' +
      added +
      ' 件\n' +
      'スキップ（条件不一致・空行等）: ' +
      skipped +
      ' 件\n' +
      'エラー（11_案件候補.エラー内容に記録）: ' +
      errors +
      ' 件';
    Logger.log(summary);
    ui.alert(summary);
  } catch (e2) {
    Logger.log(e2.stack || String(e2.message));
    ui.alert('エラー: ' + (e2.message || String(e2)));
  }
}

/**
 * @param {!Array<*>} headerRow
 */
function candidateToDispatch_validateCandidateHeaders_(headerRow) {
  var idx = candidateToDispatch_headerNameToIndex_(headerRow);
  for (var i = 0; i < candidateToDispatch_REQUIRED_CANDIDATE_HEADERS.length; i++) {
    var name = candidateToDispatch_REQUIRED_CANDIDATE_HEADERS[i];
    if (idx[name] == null) {
      throw new Error(
        candidateToDispatch_CANDIDATE_SHEET_NAME +
          ' に必須列「' +
          name +
          '」がありません（1行目のヘッダーを確認してください）。'
      );
    }
  }
}

/**
 * @param {!Array<*>} headerRow
 * @return {!Object<string, number>} 列名 → 0始まりインデックス
 */
function candidateToDispatch_headerNameToIndex_(headerRow) {
  var map = {};
  for (var c = 0; c < headerRow.length; c++) {
    var hn = normalizeCellValue_(headerRow[c]);
    if (hn) map[hn] = c;
  }
  return map;
}

/**
 * @param {!Array<*>} row
 * @param {!Object<string, number>} candCol
 * @return {boolean}
 */
function candidateToDispatch_rowMatchesTransferFilter_(row, candCol) {
  var st = normalizeCellValue_(row[candCol['案件ステータス']]);
  var cf = normalizeCellValue_(row[candCol['確認ステータス']]);
  var did = normalizeCellValue_(row[candCol['配車予定ID']]);

  if (st !== candidateToDispatch_STATUS_DECIDED) return false;
  if (cf !== candidateToDispatch_CONFIRM_DONE) return false;
  if (did !== '') return false;
  return true;
}

/**
 * スプレッドシートのタイムゾーン（取得できなければスクリプトTZ）。
 * @return {string}
 */
function candidateToDispatch_getSpreadsheetTimeZone_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = ss && ss.getSpreadsheetTimeZone();
    if (ssTz) return ssTz;
  } catch (ignore) {}
  return Session.getScriptTimeZone();
}

/**
 * 作成日時／更新日時用 IANA。スプレッドシート設定優先、無ければ Asia/Tokyo。
 * @return {string}
 */
function candidateToDispatch_getTimestampZone_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = ss && ss.getSpreadsheetTimeZone();
    if (ssTz) return ssTz;
  } catch (ignore) {}
  return 'Asia/Tokyo';
}

/**
 * 00_設定 の「配車表表示年月」に対応する年を返す（dispatch_view と同じ解釈）。
 * シート欠如・未設定・不正時はスプレッドシートTZでの実行日時の年。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} tz IANA
 * @return {number}
 */
function candidateToDispatch_getFallbackYear_(ss, tz) {
  var ym = dispatchView_readTargetYearMonthFromSettings_(ss, tz);
  return ym.year;
}

/**
 * @param {!Date} date
 * @param {string} tz IANA
 * @return {string}
 */
function candidateToDispatch_formatTimestamp_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy/MM/dd HH:mm:ss');
}

/**
 * 実行時刻を yyyy/MM/dd HH:mm:ss（スプレッドシート TZ：getTimestampZone_）で返す。
 * @return {string}
 */
function candidateToDispatch_nowText_() {
  return candidateToDispatch_formatTimestamp_(new Date(), candidateToDispatch_getTimestampZone_());
}

/**
 * 全角数字・区切りを半角に寄せる（表示崩れ対策）。
 * @param {string} s
 * @return {string}
 */
function candidateToDispatch_toAsciiDigits_(s) {
  var out = String(s)
    .replace(/\u00a0/g, ' ')
    .replace(/[\u200b\uFEFF]/g, '');
  out = out.replace(/[０-９]/g, function (ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xfee0);
  });
  return out.replace(/／/g, '/').replace(/－|―/g, '-').replace(/．/g, '.');
}

/**
 * デバッグ用にセル実体を短く文字化。
 * @param {*} v
 * @return {string}
 */
function candidateToDispatch_debugCell_(v) {
  if (v === null || v === undefined) return '(empty)';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return 'InvalidDate';
    return (
      'Date(' +
      v.getFullYear() +
      '/' +
      (v.getMonth() + 1) +
      '/' +
      v.getDate() +
      ' ' +
      v.getHours() +
      ':' +
      v.getMinutes() +
      ')'
    );
  }
  if (typeof v === 'number' && !isNaN(v)) return String(v);
  if (typeof v === 'boolean') return String(v);
  if (typeof v === 'object' && v !== null) {
    try {
      var j = JSON.stringify(v);
      if (j.length > 120) return j.substring(0, 120) + '…';
      return j;
    } catch (e) {
      return String(v);
    }
  }
  var s = normalizeCellValue_(String(v));
  if (s.length > 120) return s.substring(0, 120) + '…';
  return s || '(empty)';
}

/**
 * getDisplayValues / getValues のどちらかで暦日が取れればよい。
 * @param {*} displayVal
 * @param {*} rawVal
 * @param {string} tz
 * @param {number} fallbackYear 年なし月日の補完に使う年（00_設定に準拠）
 * @return {?{y: number, m: number, d: number}}
 */
function candidateToDispatch_parseDateTripletFromAny_(displayVal, rawVal, tz, fallbackYear) {
  var a = candidateToDispatch_parseDateOnlyParts_(displayVal, tz, fallbackYear);
  if (a) return a;
  return candidateToDispatch_parseDateOnlyParts_(rawVal, tz, fallbackYear);
}

/**
 * @param {number} sheetRowNum 1起点
 * @param {!Array<*>} row
 * @param {!Array<string>} rowDisplay
 * @param {!Object<string, number>} candCol
 * @return {string}
 */
function candidateToDispatch_formatDateParseError_(sheetRowNum, row, rowDisplay, candCol) {
  var sz = candCol['積込日'];
  var iz = candCol['納品日'];
  var id = '';
  if (candCol['案件候補ID'] != null && candCol['案件候補ID'] < row.length) {
    id = candidateToDispatch_debugCell_(row[candCol['案件候補ID']]);
  }
  function dispAt(idx) {
    if (idx == null) return '(no column)';
    if (idx >= rowDisplay.length) return '(short display row)';
    return candidateToDispatch_debugCell_(rowDisplay[idx]);
  }
  function rawAt(idx) {
    if (idx == null) return '(no column)';
    if (idx >= row.length) return '(short value row)';
    return candidateToDispatch_debugCell_(row[idx]);
  }
  return (
    '日付を解釈できません。row=' +
    sheetRowNum +
    ', 案件候補ID=' +
    id +
    ', 積込日(display)=' +
    dispAt(sz) +
    ', 積込日(value)=' +
    rawAt(sz) +
    ', 納品日(display)=' +
    dispAt(iz) +
    ', 納品日(value)=' +
    rawAt(iz)
  );
}

/**
 * カレンダー日 {y,m,d}（月は 1〜12）。解釈不能なら null。
 * 年なしの M/D・M月D日 は fallbackYear で補完。Date／シリアル／Date.parse が 2001 年に落ちる誤解釈は補完年へ差し替え。
 * @param {*} dateRaw
 * @param {string} tz
 * @param {number} fallbackYear 00_設定.配車表表示年月に連動
 * @return {?{y: number, m: number, d: number}}
 */
function candidateToDispatch_parseDateOnlyParts_(dateRaw, tz, fallbackYear) {
  if (dateRaw === null || dateRaw === undefined || dateRaw === '') return null;

  var y;
  var mo;
  var day;

  if (dateRaw instanceof Date) {
    if (isNaN(dateRaw.getTime())) return null;
    y = dateRaw.getFullYear();
    mo = dateRaw.getMonth() + 1;
    day = dateRaw.getDate();
    if (fallbackYear != null && y === 2001 && candidateToDispatch_isValidYmd_(fallbackYear, mo, day)) {
      y = fallbackYear;
    }
    if (!candidateToDispatch_isValidYmd_(y, mo, day)) return null;
    return { y: y, m: mo, d: day };
  }

  if (typeof dateRaw === 'number' && !isNaN(dateRaw)) {
    if (dateRaw > 20000) {
      var epoch = new Date(1899, 11, 30);
      var ms = epoch.getTime() + Math.round(dateRaw * 86400000);
      var dt = new Date(ms);
      if (isNaN(dt.getTime())) return null;
      y = dt.getFullYear();
      mo = dt.getMonth() + 1;
      day = dt.getDate();
      if (fallbackYear != null && y === 2001 && candidateToDispatch_isValidYmd_(fallbackYear, mo, day)) {
        y = fallbackYear;
      }
      if (!candidateToDispatch_isValidYmd_(y, mo, day)) return null;
      return { y: y, m: mo, d: day };
    }
    return null;
  }

  var text = normalizeCellValue_(String(dateRaw));
  if (!text) return null;

  var trip = candidateToDispatch_tripletFromFlexibleDateText_(text, fallbackYear);
  if (trip) return trip;

  var serialM = text.match(/^(\d{5,7})$/);
  if (serialM) {
    var sn = parseInt(serialM[1], 10);
    if (sn > 20000) return candidateToDispatch_parseDateOnlyParts_(sn, tz, fallbackYear);
  }

  var coerced = resolveScheduleDate_(dateRaw, null, tz);
  if (coerced && !isNaN(coerced.getTime())) {
    y = coerced.getFullYear();
    mo = coerced.getMonth() + 1;
    day = coerced.getDate();
    if (fallbackYear != null && y === 2001 && candidateToDispatch_isValidYmd_(fallbackYear, mo, day)) {
      y = fallbackYear;
    }
    if (candidateToDispatch_isValidYmd_(y, mo, day)) return { y: y, m: mo, d: day };
  }

  return null;
}

/**
 * 表示文字列の日付部分（先頭一致）。2026/5/26・2026-05-26・2026年5月26日。
 * fallbackYear ありのとき 5/11・5-11・5月11日（年なし）を M/D として補完。
 * @param {*} text
 * @param {number} fallbackYear
 * @return {?{y: number, m: number, d: number}}
 */
function candidateToDispatch_tripletFromFlexibleDateText_(text, fallbackYear) {
  var t = normalizeCellValue_(String(text));
  if (!t) return null;
  t = candidateToDispatch_toAsciiDigits_(t);

  var y;
  var mo;
  var day;

  var wa = t.match(/^(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})/);
  if (wa) {
    y = parseInt(wa[1], 10);
    mo = parseInt(wa[2], 10);
    day = parseInt(wa[3], 10);
    if (candidateToDispatch_isValidYmd_(y, mo, day)) return { y: y, m: mo, d: day };
  }

  var jp = t.match(/^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
  if (jp) {
    y = parseInt(jp[1], 10);
    mo = parseInt(jp[2], 10);
    day = parseInt(jp[3], 10);
    if (candidateToDispatch_isValidYmd_(y, mo, day)) return { y: y, m: mo, d: day };
  }

  if (fallbackYear != null && candidateToDispatch_isValidYmd_(fallbackYear, 1, 1)) {
    var md = t.match(/^(\d{1,2})[\/\.\-](\d{1,2})$/);
    if (md) {
      mo = parseInt(md[1], 10);
      day = parseInt(md[2], 10);
      if (candidateToDispatch_isValidYmd_(fallbackYear, mo, day)) return { y: fallbackYear, m: mo, d: day };
    }
    var mdJp = t.match(/^(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
    if (mdJp) {
      mo = parseInt(mdJp[1], 10);
      day = parseInt(mdJp[2], 10);
      if (candidateToDispatch_isValidYmd_(fallbackYear, mo, day)) return { y: fallbackYear, m: mo, d: day };
    }
  }

  return null;
}

/**
 * @param {number} y
 * @param {number} m 1〜12
 * @param {number} d
 * @return {boolean}
 */
function candidateToDispatch_isValidYmd_(y, m, d) {
  if (!y || m < 1 || m > 12 || d < 1 || d > 31) return false;
  var test = new Date(y, m - 1, d, 12, 0, 0);
  return test.getFullYear() === y && test.getMonth() === m - 1 && test.getDate() === d;
}

/**
 * @param {number} n
 * @return {string}
 */
function candidateToDispatch_pad2_(n) {
  var s = String(Math.floor(Number(n)));
  return s.length >= 2 ? s : ('0' + s).slice(-2);
}

/**
 * 表示文字列から時刻を得る（HH:mm 優先）。空・解釈不能は null。
 * @param {*} displayStr
 * @return {?{h: number, mi: number, s: number}}
 */
function candidateToDispatch_parseTimePartsFromDisplay_(displayStr) {
  if (displayStr === null || displayStr === undefined) return null;
  var text = normalizeCellValue_(String(displayStr));
  if (!text) return null;

  var tail = text;
  var combined = text.match(
    /^(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/
  );
  if (combined) {
    tail = combined[4] + ':' + combined[5] + (combined[6] ? ':' + combined[6] : '');
  }

  var ampm = tail.match(
    /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM|am|pm|午前|午後)$/i
  );
  if (ampm) {
    var h0 = parseInt(ampm[1], 10);
    var mi0 = parseInt(ampm[2], 10);
    var s0 = ampm[3] ? parseInt(ampm[3], 10) : 0;
    var apRaw = ampm[4];
    if (mi0 < 0 || mi0 > 59 || s0 < 0 || s0 > 59) return null;
    var isPm = apRaw === 'PM' || apRaw === 'pm' || apRaw === '午後';
    var isAm = apRaw === 'AM' || apRaw === 'am' || apRaw === '午前';
    if (isPm) {
      if (h0 < 1 || h0 > 12) return null;
      var h24 = h0 === 12 ? 12 : h0 + 12;
      return { h: h24, mi: mi0, s: s0 };
    }
    if (isAm) {
      if (h0 < 1 || h0 > 12) return null;
      var h24a = h0 === 12 ? 0 : h0;
      return { h: h24a, mi: mi0, s: s0 };
    }
  }

  var m24 = tail.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (m24) {
    var h = parseInt(m24[1], 10);
    var mi = parseInt(m24[2], 10);
    var s = m24[3] ? parseInt(m24[3], 10) : 0;
    if (h >= 0 && h <= 23 && mi >= 0 && mi <= 59 && s >= 0 && s <= 59) {
      return { h: h, mi: mi, s: s };
    }
  }
  return null;
}

/**
 * 拡張列用。日時は **yyyy/MM/dd HH:mm** の文字列。時刻未入力なら ''。
 * 日付は表示値を優先し、壊れている場合のみ getValues() をフォールバック。
 * @param {!Array<*>} row
 * @param {!Array<string>} rowDisplay
 * @param {!Object<string, number>} candCol
 * @param {string} dateColName
 * @param {string} timeColName
 * @param {string} tz
 * @param {number} fallbackYear
 * @return {string}
 */
function candidateToDispatch_buildDatetimeYmdHmDisplay_(
  row,
  rowDisplay,
  candCol,
  dateColName,
  timeColName,
  tz,
  fallbackYear
) {
  var dateIdx = candCol[dateColName];
  if (dateIdx == null) return '';

  var dateDisplay = dateIdx < rowDisplay.length ? rowDisplay[dateIdx] : '';
  var dateValue = dateIdx < row.length ? row[dateIdx] : '';
  var p = candidateToDispatch_parseDateTripletFromAny_(dateDisplay, dateValue, tz, fallbackYear);
  if (!p) return '';

  var timeIdx = candCol[timeColName];
  if (timeIdx == null) return '';

  var timeDisplay = timeIdx < rowDisplay.length ? rowDisplay[timeIdx] : '';
  var timeValue = timeIdx < row.length ? row[timeIdx] : '';
  var tp = candidateToDispatch_parseTimePartsFromDisplay_(timeDisplay);
  if (!tp) tp = candidateToDispatch_parseTimePartsFromInput_(timeValue);
  if (!tp) return '';

  return (
    p.y +
    '/' +
    candidateToDispatch_pad2_(p.m) +
    '/' +
    candidateToDispatch_pad2_(p.d) +
    ' ' +
    candidateToDispatch_pad2_(tp.h) +
    ':' +
    candidateToDispatch_pad2_(tp.mi)
  );
}

/**
 * 時刻セルから {h,mi,s} を得る。未入力・解釈不能なら null。
 * 「時刻のみ」セルは 1899-12-30 基準の Date になる。**0:0:0 のみ**未入力扱い。
 * 同じ日に 16:00 等が入っている場合は本物の時刻として採用する。
 * @param {*} timeRaw
 * @return {?{h: number, mi: number, s: number}}
 */
function candidateToDispatch_parseTimePartsFromInput_(timeRaw) {
  if (timeRaw === null || timeRaw === undefined || timeRaw === '') return null;

  if (timeRaw instanceof Date) {
    if (isNaN(timeRaw.getTime())) return null;
    var y = timeRaw.getFullYear();
    var mo = timeRaw.getMonth();
    var da = timeRaw.getDate();
    var h = timeRaw.getHours();
    var mi = timeRaw.getMinutes();
    var s = timeRaw.getSeconds();
    var isEpochPlaceholder = y === 1899 && mo === 11 && da === 30;
    if (isEpochPlaceholder && h === 0 && mi === 0 && s === 0) return null;
    return { h: h, mi: mi, s: s };
  }

  if (typeof timeRaw === 'number' && !isNaN(timeRaw)) {
    if (timeRaw === 0) return null;
    if (timeRaw > 0 && timeRaw < 1) {
      var ms = Math.round(timeRaw * 86400000);
      return {
        h: Math.floor(ms / 3600000) % 24,
        mi: Math.floor((ms % 3600000) / 60000),
        s: Math.floor((ms % 60000) / 1000),
      };
    }
    return null;
  }

  var text = normalizeCellValue_(String(timeRaw));
  if (!text) return null;
  var m = text.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (m) {
    return {
      h: parseInt(m[1], 10),
      mi: parseInt(m[2], 10),
      s: m[3] ? parseInt(m[3], 10) : 0,
    };
  }
  return null;
}

/**
 * 予定日は積込日を優先し、無ければ納品日。表示値・実値のどちらかで暦日を決める。
 * @param {!Array<*>} row
 * @param {!Array<string>} rowDisplay
 * @param {!Object<string, number>} candCol
 * @param {number} sheetRowNum 1起点
 * @param {number} fallbackYear 年なし月日の補完
 * @return {!Date}
 */
function candidateToDispatch_resolvePlanDate_(row, rowDisplay, candCol, sheetRowNum, fallbackYear) {
  var pickupIdx = candCol['積込日'];
  var deliveryIdx = candCol['納品日'];
  var tz = candidateToDispatch_getSpreadsheetTimeZone_();

  if (pickupIdx != null) {
    var pickupDateDisplay = pickupIdx < rowDisplay.length ? rowDisplay[pickupIdx] : '';
    var pickupDateValue = pickupIdx < row.length ? row[pickupIdx] : '';
    var p1 = candidateToDispatch_parseDateTripletFromAny_(pickupDateDisplay, pickupDateValue, tz, fallbackYear);
    if (p1) return new Date(p1.y, p1.m - 1, p1.d, 12, 0, 0);
  }
  if (deliveryIdx != null) {
    var deliveryDateDisplay = deliveryIdx < rowDisplay.length ? rowDisplay[deliveryIdx] : '';
    var deliveryDateValue = deliveryIdx < row.length ? row[deliveryIdx] : '';
    var p2 = candidateToDispatch_parseDateTripletFromAny_(deliveryDateDisplay, deliveryDateValue, tz, fallbackYear);
    if (p2) return new Date(p2.y, p2.m - 1, p2.d, 12, 0, 0);
  }
  throw new Error(candidateToDispatch_formatDateParseError_(sheetRowNum, row, rowDisplay, candCol));
}

/**
 * @param {!Array<*>} row
 * @param {!Array<string>} rowDisplay
 * @param {!Object<string, number>} candCol
 * @param {number} sheetRowNum 1起点
 * @param {string} timestampText yyyy/MM/dd HH:mm:ss（スプレッドシート TZ）
 * @param {number} fallbackYear 00_設定.配車表表示年月から求める補完年
 * @return {!Object<string, *>}
 */
function candidateToDispatch_buildMergedFromCandidateRow_(
  row,
  rowDisplay,
  candCol,
  sheetRowNum,
  timestampText,
  fallbackYear
) {
  var planDate = candidateToDispatch_resolvePlanDate_(row, rowDisplay, candCol, sheetRowNum, fallbackYear);
  var y = planDate.getFullYear();
  var m = planDate.getMonth() + 1;
  var d = planDate.getDate();
  var planText = buildScheduleDateText_(y, m, d);
  var weekday = getWeekdayText_(planDate);

  var jobName = candidateToDispatch_pickOptionalStringPreferDisplay_(row, rowDisplay, candCol, '案件名');
  if (!jobName) throw new Error('案件名が空です。');

  var provCol1 = candidateToDispatch_provenanceColumn1_(candCol);
  var cellA1 = toA1_(sheetRowNum, provCol1);

  var pickupLoc = candidateToDispatch_pickOptionalStringPreferDisplay_(
    row,
    rowDisplay,
    candCol,
    '積込地'
  );
  var dropLoc = candidateToDispatch_pickOptionalStringPreferDisplay_(row, rowDisplay, candCol, '納品地');

  /** @type {!Object<string, *>} */
  var merged = {};

  merged['運行予定ID'] = createScheduleId_();
  merged['運行予定番号'] = '';
  merged['予定日'] = planText;
  merged['曜日'] = weekday;
  merged['案件ID'] = '';
  merged['案件名_入力値'] = jobName;
  merged['運転者ID'] = '';
  merged['運転者名_入力値'] = '';
  merged['車両ID'] = '';
  merged['車両呼称_入力値'] = '';

  merged['ステータス'] = INITIAL_STATUS;
  // 候補反映では出発・到着「予定時刻」列には書かない。
  merged['出発地'] = pickupLoc;
  merged['目的地'] = dropLoc;
  merged['積込地'] = pickupLoc;
  merged['納品地'] = dropLoc;

  merged['配車ステータス'] = candidateToDispatch_SCHEDULE_ASSIGN_STATUS;

  var tz = candidateToDispatch_getSpreadsheetTimeZone_();
  merged['積込予定日時'] = candidateToDispatch_buildDatetimeYmdHmDisplay_(
    row,
    rowDisplay,
    candCol,
    '積込日',
    '積込時間',
    tz,
    fallbackYear
  );
  merged['納品予定日時'] = candidateToDispatch_buildDatetimeYmdHmDisplay_(
    row,
    rowDisplay,
    candCol,
    '納品日',
    '納品時間',
    tz,
    fallbackYear
  );

  merged['作業内容'] = '';
  merged['備考'] = candidateToDispatch_pickOptionalStringPreferDisplay_(row, rowDisplay, candCol, '補足メモ');
  merged['摘要・ルート等'] = '';
  merged['カレンダー表示名'] = '';
  merged['カレンダーイベントID'] = '';
  merged['最終同期日時'] = '';
  merged['LINE送信済みフラグ'] = false;
  merged['取引先ID'] = '';

  merged['作成元シート名'] = candidateToDispatch_CANDIDATE_SHEET_NAME;
  merged['作成元セル'] = cellA1;
  merged['マスタ照合状態'] = MASTER_PENDING;
  merged['照合エラーメモ'] = '';
  merged['作成日時'] = timestampText;
  merged['更新日時'] = timestampText;

  candidateToDispatch_copyOptionalSameNameFields_(merged, row, rowDisplay, candCol, [
    '案件候補ID',
    '荷主名',
    '案件名',
    '荷種',
    '重量',
    '車格条件',
    '台数条件',
    '運賃',
    '高速代',
  ]);

  return merged;
}

/**
 * @param {!Object<string, number>} candCol
 * @return {number} 1起点の列番号（作成元セル用）
 */
function candidateToDispatch_provenanceColumn1_(candCol) {
  if (candCol['案件候補ID'] != null) return candCol['案件候補ID'] + 1;
  if (candCol['案件名'] != null) return candCol['案件名'] + 1;
  return 1;
}

/**
 * 表示値があれば優先（地名・メモ・日時の見た目と一致）。
 * @param {!Array<*>} row
 * @param {!Array<string>} rowDisplay
 * @param {!Object<string, number>} candCol
 * @param {string} name
 * @return {string}
 */
function candidateToDispatch_pickOptionalStringPreferDisplay_(row, rowDisplay, candCol, name) {
  if (candCol[name] == null) return '';
  var idx = candCol[name];
  if (idx < rowDisplay.length) {
    var disp = rowDisplay[idx];
    if (disp !== null && disp !== undefined && normalizeCellValue_(String(disp)) !== '') {
      return normalizeCellValue_(String(disp));
    }
  }
  return normalizeCellValue_(row[idx]);
}

/**
 * 10 のヘッダに同名の列がある場合だけ候補値を転記する（表示値優先）。
 * @param {!Object<string, *>} merged
 * @param {!Array<*>} row
 * @param {!Array<string>} rowDisplay
 * @param {!Object<string, number>} candCol
 * @param {!Array<string>} names
 */
function candidateToDispatch_copyOptionalSameNameFields_(merged, row, rowDisplay, candCol, names) {
  for (var i = 0; i < names.length; i++) {
    var nm = names[i];
    if (candCol[nm] == null) continue;
    var idx = candCol[nm];
    var useDisplay =
      idx < rowDisplay.length &&
      rowDisplay[idx] !== null &&
      rowDisplay[idx] !== undefined &&
      normalizeCellValue_(String(rowDisplay[idx])) !== '';
    merged[nm] = useDisplay ? rowDisplay[idx] : row[idx];
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} scheduleSheet
 * @param {!Array<*>} headerRowVals
 * @param {!Object<string, *>} merged
 */
function candidateToDispatch_appendScheduleRow_(scheduleSheet, headerRowVals, merged) {
  var numColsPhysical = headerRowVals.length;
  var line = buildOutputLine_(headerRowVals, merged);
  var sanitized = sanitizeLineForSet_(line, numColsPhysical);
  if (!sanitized || isBlankPhysicalLine_(sanitized)) {
    throw new Error('10_配車予定への出力行が空になりました。');
  }
  var startRow = scheduleSheet.getLastRow() + 1;
  scheduleSheet.getRange(startRow, 1, 1, numColsPhysical).setValues([sanitized]);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} candidateSheet
 * @param {number} sheetRowNum
 * @param {!Object<string, number>} candCol
 * @param {string} scheduleId
 * @param {string} updateTimestampText yyyy/MM/dd HH:mm:ss
 */
function candidateToDispatch_patchCandidateAfterSuccess_(
  candidateSheet,
  sheetRowNum,
  candCol,
  scheduleId,
  updateTimestampText
) {
  if (candCol['配車予定ID'] != null) {
    candidateSheet.getRange(sheetRowNum, candCol['配車予定ID'] + 1).setValue(scheduleId);
  }
  if (candCol['案件ステータス'] != null) {
    candidateSheet
      .getRange(sheetRowNum, candCol['案件ステータス'] + 1)
      .setValue(candidateToDispatch_STATUS_DISPATCHED);
  }
  if (candCol['更新日時'] != null) {
    candidateSheet.getRange(sheetRowNum, candCol['更新日時'] + 1).setValue(updateTimestampText);
  }
  if (candCol['エラー内容'] != null) {
    candidateSheet.getRange(sheetRowNum, candCol['エラー内容'] + 1).setValue('');
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} candidateSheet
 * @param {number} sheetRowNum
 * @param {!Object<string, number>} candCol
 * @param {string} message
 */
function candidateToDispatch_writeCandidateError_(candidateSheet, sheetRowNum, candCol, message) {
  if (candCol['エラー内容'] == null) return;
  candidateSheet.getRange(sheetRowNum, candCol['エラー内容'] + 1).setValue(message);
}
