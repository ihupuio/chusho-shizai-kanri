// ===== 資材報告 受信サーバー (Google Apps Script) =====
//
// 使い方：
// 1. Googleスプレッドシートを新規作成
// 2. 拡張機能 > Apps Script を開き、このファイルの中身を貼り付けて保存
// 3. エディタ上部の関数選択で「setup」を選んで実行（初回は権限承認あり）
//    → 「記録」「最新」シートが自動で作られる
// 4. デプロイ > 新しいデプロイ > 種類：ウェブアプリ
//    - 次のユーザーとして実行：自分
//    - アクセスできるユーザー：全員
//    → 発行されたURL(…/exec)を docs/config.js の GAS_URL に貼る
//
// シート構成：
//   記録 … 全報告の生データ（1報告 = 資材ごとに複数行、報告IDで束ねる）
//   最新 … 車両ごとの最新報告のマトリクス（Googleサイト埋め込み用・自動更新）

const SHEET_RECORD = "記録";
const SHEET_LATEST = "最新";

// 初期セットアップ：シートを作る
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let rec = ss.getSheetByName(SHEET_RECORD);
  if (!rec) {
    rec = ss.insertSheet(SHEET_RECORD);
    rec.getRange(1, 1, 1, 8).setValues([
      ["日時", "報告ID", "氏名", "LINE UserID", "車両", "資材", "数量", "備考"],
    ]);
    rec.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SHEET_LATEST)) {
    const latest = ss.insertSheet(SHEET_LATEST);
    latest.getRange(1, 1).setValue("（報告が届くとここに車両×資材の最新表が自動生成されます）");
  }
}

// LIFFフォームからの送信を受け取る
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (!data.vehicle) return jsonOutput({ ok: false, error: "車両が未指定です" });
    if (!data.items || typeof data.items !== "object" || Object.keys(data.items).length === 0) {
      return jsonOutput({ ok: false, error: "資材データがありません" });
    }
    for (const k in data.items) {
      const v = Number(data.items[k]);
      if (isNaN(v) || v < 0) return jsonOutput({ ok: false, error: "数量が不正です：" + k });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const now = new Date();
      const reportId =
        Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMdd-HHmmss") +
        "-" + String(data.userId || "x").slice(-6);

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORD);
      const rows = Object.keys(data.items).map((item) => [
        now,
        reportId,
        String(data.name || ""),
        String(data.userId || ""),
        String(data.vehicle),
        item,
        Number(data.items[item]),
        String(data.note || ""),
      ]);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);

      updateLatestSheet();
    } finally {
      lock.releaseLock();
    }

    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

// ダッシュボード用：車両ごとの最新報告をJSONで返す
function doGet() {
  try {
    return jsonOutput({ ok: true, vehicles: getLatestReports() });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

// 記録シートから車両ごとの最新報告を組み立てる
function getLatestReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORD);
  const values = sheet.getDataRange().getValues().slice(1);
  const byVehicle = {};

  values.forEach((r) => {
    const [time, reportId, name, userId, vehicle, item, qty, note] = r;
    if (!vehicle || !item) return;
    const cur = byVehicle[vehicle];
    if (!cur || cur.reportId !== reportId) {
      // 新しい報告IDの行が来たら、より新しい場合のみ置き換える
      if (cur && new Date(time) < new Date(cur.time)) return;
      byVehicle[vehicle] = {
        vehicle: String(vehicle),
        reportId: reportId,
        time: new Date(time).toISOString(),
        name: String(name || ""),
        note: String(note || ""),
        items: {},
      };
    }
    byVehicle[vehicle].items[String(item)] = Number(qty);
  });

  return Object.keys(byVehicle).sort().map((v) => byVehicle[v]);
}

// 「最新」シートを車両×資材のマトリクスに書き直す（Googleサイト埋め込み用）
function updateLatestSheet() {
  const reports = getLatestReports();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LATEST);
  sheet.clearContents();
  if (reports.length === 0) return;

  // 資材名の一覧（登場順）
  const itemNames = [];
  reports.forEach((r) => {
    Object.keys(r.items).forEach((n) => { if (!itemNames.includes(n)) itemNames.push(n); });
  });

  const header = ["車両", "報告日時", "報告者"].concat(itemNames, ["備考"]);
  const rows = reports.map((r) =>
    [r.vehicle, new Date(r.time), r.name]
      .concat(itemNames.map((n) => (n in r.items ? r.items[n] : "")), [r.note])
  );

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
