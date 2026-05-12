/**
 * 10_配車予定：未配車かつ運転者ID・車両IDが手入力済みの行の配車ステータスを「配車済」に更新する。
 * dispatch_conversion.gs / candidate_to_dispatch.gs とは独立。ヘルパーは dispatchAssignment_ 接頭辞。
 *
 * @fileoverview 運転者名_入力値・車両呼称_入力値のマスタ補完は行わない（ID手入力後のステータス整合のみ）。
 */

/** @type {string} */
var dispatchAssignment_STATUS_UNASSIGNED = '未配車';

/** @type {string} */
var dispatchAssignment_STATUS_ASSIGNED = '配車済';

/** @type {!Array<string>} 本処理で必須の 10_配車予定 ヘッダー（1行目） */
var dispatchAssignment_REQUIRED_HEADERS = [
  '配車ステータス',
  '運転者ID',
  '車両ID',
  '更新日時',
];

/**
 * メニューから実行。
 * 条件: 配車ステータス=未配車 かつ 運転者ID・車両ID がともに非空。
 */
function updateAssignedDispatchStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    var sheet = ss.getSheetByName(TARGET_SCHEDULE_SHEET_NAME);
    if (!sheet) {
      throw new Error('シート「' + TARGET_SCHEDULE_SHEET_NAME + '」が見つかりません。');
    }

    var headerMap = getSheetHeaderMap_(sheet);
    dispatchAssignment_validateHeadersForUpdate_(headerMap);

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      ui.alert('対象データがありません（' + TARGET_SCHEDULE_SHEET_NAME + '）。');
      return;
    }

    var tz = dispatchAssignment_getSpreadsheetTimeZone_();
    var timestampText = dispatchAssignment_formatNow_(tz);

    var colStatus = headerMap['配車ステータス'];
    var colDriverId = headerMap['運転者ID'];
    var colVehicleId = headerMap['車両ID'];
    var colUpdated = headerMap['更新日時'];

    var dataRange = sheet.getRange(2, 1, lastRow, lastCol);
    var rows = dataRange.getValues();

    var updated = 0;
    var i;
    var r;
    var st;
    var did;
    var vid;

    for (i = 0; i < rows.length; i++) {
      r = rows[i];
      st = normalizeCellValue_(r[colStatus - 1]);
      did = normalizeCellValue_(r[colDriverId - 1]);
      vid = normalizeCellValue_(r[colVehicleId - 1]);

      if (st !== dispatchAssignment_STATUS_UNASSIGNED) continue;
      if (did === '' || vid === '') continue;

      r[colStatus - 1] = dispatchAssignment_STATUS_ASSIGNED;
      r[colUpdated - 1] = timestampText;
      updated++;
    }

    if (updated > 0) {
      dataRange.setValues(rows);
    }

    var msg =
      '割当状態の更新完了\n' +
      '更新した行数: ' +
      updated +
      ' 件\n' +
      '（配車ステータス=未配車 かつ 運転者ID・車両ID 両方あり のみ）';
    Logger.log(msg);
    ui.alert(msg);
  } catch (e) {
    Logger.log(e.stack || String(e.message));
    ui.alert('エラー: ' + (e.message || String(e)));
  }
}

/**
 * @param {!Object<string, number>} headerMap
 */
function dispatchAssignment_validateHeadersForUpdate_(headerMap) {
  var i;
  var name;
  for (i = 0; i < dispatchAssignment_REQUIRED_HEADERS.length; i++) {
    name = dispatchAssignment_REQUIRED_HEADERS[i];
    if (!headerMap[name]) {
      throw new Error(
        TARGET_SCHEDULE_SHEET_NAME +
          ' に必須列「' +
          name +
          '」がありません（1行目のヘッダーを確認してください）。'
      );
    }
  }
}

/**
 * @return {string}
 */
function dispatchAssignment_getSpreadsheetTimeZone_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = ss && ss.getSpreadsheetTimeZone();
    if (ssTz) return ssTz;
  } catch (ignore) {}
  return Session.getScriptTimeZone() || 'Asia/Tokyo';
}

/**
 * @param {string} tz IANA
 * @return {string}
 */
function dispatchAssignment_formatNow_(tz) {
  return Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');
}
