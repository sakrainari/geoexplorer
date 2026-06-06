/**
 * GeoExplorer — スプレッドシート書き込み用 Google Apps Script
 *
 * 役割：
 *  - index.html（入力ツール）から送られてきたデータを受け取り、
 *    「スポットマスター」と「解説シート」へ同時に1行ずつ追加する。
 *  - IDは自動採番し、両シートに同じIDを書き込む。
 *
 * デプロイ：このスクリプトを対象スプレッドシートのコンテナバインドとして作成し、
 *          「ウェブアプリ」としてデプロイする（手順はリポジトリの DEPLOY.md 参照）。
 */

// ===== 設定 =====
var MASTER_SHEET = 'スポットマスター';  // スポットマスターのシート名
var DESC_SHEET   = '解説シート';        // 解説シートのシート名
var ID_PREFIX    = 'SPOT-';             // IDの接頭辞（不要なら '' にする）
var ID_PAD       = 4;                   // ID連番のゼロ埋め桁数（例：4 → 0001）

/**
 * ヘルスチェック用。デプロイ後にブラウザでURLを開くと
 * シートの状況が表示され、接続できているか確認できる。
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(MASTER_SHEET);
  var desc   = ss.getSheetByName(DESC_SHEET);
  var info = {
    ok: !!(master && desc),
    spreadsheet: ss.getName(),
    masterSheet: master ? (MASTER_SHEET + '（最終行:' + master.getLastRow() + '）') : '見つかりません',
    descSheet:   desc   ? (DESC_SHEET   + '（最終行:' + desc.getLastRow()   + '）') : '見つかりません',
    nextId: master ? nextId_(master) : null
  };
  return json_(info);
}

/**
 * 入力ツールからのPOSTを受け取り、2シートへ1行ずつ追加する。
 * 受信ボディは JSON 文字列（{ spotMaster:{...}, descSheet:{...} }）。
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 同時実行による採番衝突を防ぐ

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: '送信データが空です。' });
    }
    var data = JSON.parse(e.postData.contents);
    var m = data.spotMaster || {};
    var d = data.descSheet  || {};

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName(MASTER_SHEET);
    var desc   = ss.getSheetByName(DESC_SHEET);
    if (!master || !desc) {
      return json_({ ok: false, error: 'シートが見つかりません（必要: ' + MASTER_SHEET + ' / ' + DESC_SHEET + '）' });
    }

    // --- ID自動採番（既存IDの最大連番 + 1） ---
    var newId = nextId_(master);

    // --- スポットマスター 1行（A〜S の19列） ---
    var masterRow = [
      newId,                          // A: ID
      m.グループID        || '',       // B: グループID
      m.大陸              || '',       // C: 大陸
      m.国                || '',       // D: 国
      m.地区都市          || '',       // E: 地区・都市
      m.スポット名_日     || '',       // F: スポット名（日本語）
      m.スポット名_英     || '',       // G: スポット名（英語）
      m.SVリンク          || '',       // H: SVリンク
      m.SV種別            || '',       // I: SV種別
      m.SVリンク最終確認日 || '',       // J: SVリンク最終確認日
      m.緯度              || '',       // K: 緯度
      m.経度              || '',       // L: 経度
      toCell_(m.タグ),                 // M: タグ（配列→「、」区切り）
      toCell_(m.世界遺産登録基準),      // N: 世界遺産登録基準（配列→「、」区切り）
      m.Wikipedia画像URL  || '',       // O: Wikipedia画像URL
      m.画像クレジット    || '',       // P: 画像クレジット
      m.SV登録状況        || '',       // Q: SV登録状況
      m.公開ステータス    || '',       // R: 公開ステータス
      m.備考              || ''        // S: 備考
    ];

    // --- 解説シート 1行（A〜N の14列） ---
    var descRow = [
      newId,                          // A: ID（マスターと共有）
      d.基本解説          || '',       // B: 基本解説
      d.ピンポイント情報  || '',       // C: ピンポイント情報
      d.語源由来          || '',       // D: 語源・由来
      d.関連する人物      || '',       // E: 関連する人物
      d.人間ドラマ        || '',       // F: 人間ドラマ
      d.世界遺産検定メモ  || '',       // G: 世界遺産検定メモ
      d.登録基準の解説    || '',       // H: 登録基準の解説
      d.見どころ          || '',       // I: 見どころ・注目ポイント
      d.雑学うんちく      || '',       // J: 雑学・うんちく
      d.伝説逸話          || '',       // K: 伝説・逸話
      d.参考文献          || '',       // L: 参考文献
      d.解説品質フラグ    || '',       // M: 解説品質フラグ
      d.最終更新日        || ''        // N: 最終更新日
    ];

    // 2シートへ追記（同一トランザクション的に連続実行）
    master.appendRow(masterRow);
    desc.appendRow(descRow);
    SpreadsheetApp.flush();

    return json_({ ok: true, id: newId });

  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 次のIDを返す。スポットマスターA列の既存IDから末尾の数字の最大値を求め、+1する。
 * （行削除があっても連番が重複しない・ヘッダー行は自動的に無視される）
 */
function nextId_(master) {
  var lastRow = master.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var ids = master.getRange(2, 1, lastRow - 1, 1).getValues(); // A2:A（ヘッダー除く）
    for (var i = 0; i < ids.length; i++) {
      var s = String(ids[i][0]);
      var match = s.match(/(\d+)\s*$/);
      if (match) {
        var n = parseInt(match[1], 10);
        if (n > max) max = n;
      }
    }
  }
  return formatId_(max + 1);
}

/** 連番を ID 文字列に整形（例：1 → "SPOT-0001"） */
function formatId_(n) {
  var num = String(n);
  while (num.length < ID_PAD) num = '0' + num;
  return ID_PREFIX + num;
}

/** 配列なら「、」で結合、それ以外は文字列化してセル値にする */
function toCell_(v) {
  if (Array.isArray(v)) return v.join('、');
  return v == null ? '' : String(v);
}

/** JSONレスポンスを返す */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
