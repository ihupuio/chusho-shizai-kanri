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
const SHEET_WORKERS = "作業者マスタ";
const SHEET_VEHICLES = "車両マスタ";
const INITIAL_VEHICLES = ["3620", "4078", "4079", "4913", "4914", "5433", "1552"];

// 初期セットアップ：シートを作る
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let rec = ss.getSheetByName(SHEET_RECORD);
  if (!rec) {
    rec = ss.insertSheet(SHEET_RECORD);
    rec.getRange(1, 1, 1, 11).setValues([
      ["日時", "報告ID", "氏名", "LINE UserID", "種別", "資材", "変更数", "移動元", "移動先", "車両", "備考"],
    ]);
    rec.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SHEET_LATEST)) {
    const latest = ss.insertSheet(SHEET_LATEST);
    latest.getRange(1, 1).setValue("（報告が届くとここに車両×資材の最新表が自動生成されます）");
  }

  if (!ss.getSheetByName(SHEET_WORKERS)) {
    const workers = ss.insertSheet(SHEET_WORKERS);
    workers.getRange(1, 1, 1, 3).setValues([["氏名", "LINE表示名", "有効"]]);
    workers.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_VEHICLES)) {
    const vehicles = ss.insertSheet(SHEET_VEHICLES);
    vehicles.getRange(1, 1, 1, 3).setValues([["号車番号", "状態", "表示順"]]);
    vehicles.getRange(2, 1, INITIAL_VEHICLES.length, 3).setValues(
      INITIAL_VEHICLES.map((v, i) => [v, "使用中", i + 1])
    );
    vehicles.setFrozenRows(1);
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
      if (isNaN(v) || v < 0 || (data.mode !== "inventory" && v === 0)) return jsonOutput({ ok: false, error: "数量が不正です：" + k });
    }
    const type = data.mode === "inventory" ? "棚卸し" : ({ new: "新規", move: "移動", dispose: "廃棄" }[data.changeType] || "新規");
    if (type === "移動" && (!data.from || !data.to || data.from === data.to)) return jsonOutput({ ok: false, error: "移動元と移動先が不正です" });
    if (type === "廃棄" && !data.from) return jsonOutput({ ok: false, error: "廃棄元が未指定です" });

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
        type,
        item,
        Number(data.items[item]),
        String(data.from || (type === "新規" ? "倉庫" : "")),
        String(data.to || (type === "新規" ? data.vehicle : type === "廃棄" ? "廃棄" : "")),
        String(data.vehicle),
        String(data.note || ""),
      ]);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);

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
    return jsonOutput({ ok: true, vehicles: getLatestReports(), movements: getMovementHistory(), vehicleMaster: getVehicleMaster(), workerMaster: getWorkerMaster() });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function getVehicleMaster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VEHICLES);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter((r) => r[0] && String(r[1] || "使用中") === "使用中")
    .sort((a, b) => Number(a[2] || 0) - Number(b[2] || 0))
    .map((r) => String(r[0]));
}

function getWorkerMaster() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_WORKERS);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter((r) => r[0] && String(r[2] || "有効") !== "無効")
    .map((r) => ({ name: String(r[0]), lineName: String(r[1] || "") }));
}

function getMovementHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORD);
  return sheet.getDataRange().getValues().slice(1)
    .filter((r) => r[4] && r[4] !== "棚卸し")
    .sort((a, b) => new Date(b[0]) - new Date(a[0]))
    .slice(0, 100)
    .map((r) => ({ time: new Date(r[0]).toISOString(), type: String(r[4]), item: String(r[5]), qty: Number(r[6]), from: String(r[7] || ""), to: String(r[8] || ""), vehicle: String(r[9] || ""), name: String(r[2] || "") }));
}

// 記録シートから車両ごとの最新報告を組み立てる
function getLatestReports() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORD);
  const values = sheet.getDataRange().getValues().slice(1);
  const stock = {};
  const latest = {};
  values.sort((a, b) => new Date(a[0]) - new Date(b[0])).forEach((r) => {
    const [time, reportId, name, userId, type, item, qty, from, to, vehicle, note] = r;
    if (!item || !type) return;
    const n = Number(qty) || 0;
    const ensure = (v) => { if (v && v !== "新規" && v !== "廃棄") { stock[v] = stock[v] || {}; stock[v][item] = Number(stock[v][item] || 0); } };
    if (type === "棚卸し") { ensure(vehicle); stock[vehicle][item] = n; }
    if (type === "新規") { ensure(to || vehicle); stock[to || vehicle][item] += n; }
    if (type === "移動") { ensure(from); ensure(to); if (from && from !== "新規") stock[from][item] -= n; if (to && to !== "廃棄" && to !== "新規") stock[to][item] += n; }
    if (type === "廃棄") { ensure(from || vehicle); if ((from || vehicle) !== "倉庫") stock[from || vehicle][item] -= n; }
    const affected = [vehicle, from, to].filter((v) => v && v !== "新規" && v !== "廃棄");
    affected.forEach((v) => { latest[v] = { vehicle: v, reportId: reportId, time: new Date(time).toISOString(), name: String(name || ""), note: String(note || ""), changeType: type, from: String(from || ""), to: String(to || ""), items: Object.assign({}, stock[v]) }; });
  });
  return Object.keys(stock).sort().map((v) => Object.assign({ vehicle: v, items: stock[v] }, latest[v] || {}));
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
