// ===== 設定ファイル =====
// セットアップ時にここを書き換えるだけでOK

const CONFIG = {
  // LINE DevelopersでLIFFアプリを登録すると発行されるID（例: "2001234567-AbcdEfgh"）
  // 空のままだと「デモモード」になり、LINEなしのブラウザで動作確認できます
  LIFF_ID: "2010766899-wHZg1t0X",

  // GASをウェブアプリとしてデプロイしたときのURL（https://script.google.com/macros/s/…/exec）
  GAS_URL: "https://script.google.com/macros/s/AKfycbxI6A0NIRZDhgDf3ajrWO7MuuAYYuEcKWNpdtU57jUs11ir9b59SyRnlVvuRQ94kjWeFg/exec",

  // 車両のリスト（自由に書き換えてください）
  VEHICLES: [
    "3620",
    "4078",
    "4079",
    "4913",
    "4914",
    "5433",
    "1552",
  ],

  // 資材のリスト
  //   name: 資材名 / std: 定数（車両に積んでおくべき数。ダッシュボードで不足を色分け）
  //   定数管理をしない資材は std を 0 にすればそのまま表示のみ
  //   増やしたい場合は行を足すだけ（フォーム・ダッシュボードとも自動で対応）
  ITEMS: [
    { name: "ストレッチフィルム", std: 2 },
  ],

  // ダッシュボードの自動更新間隔（秒）
  REFRESH_SEC: 30,
};
