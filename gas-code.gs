/**
 * 移動サロン2026 クイズ選手権 集計サーバー (Google Apps Script)
 * 貼り付けたら、まず setup() を1回実行してシートを初期化してください。
 * アプリからのJSON POSTを受け取り、チェックコードを再計算・照合して
 * 「エントリー」シートに1行追記します。
 */

var CODE_SALT = "ido-salon-2026-tsukuba";  // index.html の CODE_SALT と一致必須
var ENTRY_SHEET = "エントリー";
var RANK_SHEET = "ランキング";

/* 正解キー(index.html の QUIZ の a 値と一致必須) */
var ANSWERS = {
  gsi:  [2, 0, 1, 0, 2, 1, 2],
  geo:  [1, 1, 0, 0, 0, 1],
  jaxa: [0, 0, 1, 0, 0, 1, 1]
};

/* 動作確認用: ウェブアプリURLをブラウザで開くと ok と表示される */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: "ido-salon-2026" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var out;
  try {
    var d = JSON.parse(e.postData.contents);
    var name = String(d.name || "").trim().slice(0, 40);
    var score = String(d.score || "");
    var guess = String(d.guess || "");
    var finished = String(d.finished || "");
    var pattern = String(d.pattern || "").slice(0, 80);
    var code = String(d.code || "");

    if (!name || !/^\d+\/\d+$/.test(score) || !/^\d+$/.test(guess)) {
      out = { ok: false, error: "bad-request" };
    } else {
      var m = finished.match(/\((\d+)\)/);
      var finishedMs = m ? m[1] : "";
      var codeOk = (makeCode(name, pattern, guess, finishedMs) === code);
      var scoreOk = (scoreFromPattern(pattern) === score);
      var problems = [];
      if (!codeOk) problems.push("コード不一致");
      if (!scoreOk) problems.push("スコアとパターン不一致");
      var flag = problems.length ? "⚠️ " + problems.join("・") : "OK";

      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var sh = getEntrySheet();
        /* 同名の2回目以降は🔁重複を付与(再送の保険は許容し、集計時は最初の行を採用) */
        var last = sh.getLastRow();
        if (last >= 2) {
          var names = sh.getRange(2, 2, last - 1, 1).getValues();
          for (var i = 0; i < names.length; i++) {
            if (String(names[i][0]) === name) { flag += "・🔁重複"; break; }
          }
        }
        sh.appendRow([new Date(), name, score, guess, finished, pattern, code, flag]);
      } finally {
        lock.releaseLock();
      }
      out = { ok: true, verified: problems.length === 0 };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 初期化: エントリーシートとランキングシート(自動計算式つき)を作成 */
function setup() {
  getEntrySheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r = ss.getSheetByName(RANK_SHEET) || ss.insertSheet(RANK_SHEET);
  r.getRange("A1").setValue("ニアピン正解(全員の正解数合計)→");
  r.getRange("B1").setFormula('=SUM(ARRAYFORMULA(IFERROR(VALUE(LEFT(エントリー!C2:C,FIND("/",エントリー!C2:C)-1)))))');
  r.getRange("A2:D2").setValues([["名前", "正解数", "ニアピン差", "完了時刻(ミリ秒)"]]);
  r.getRange("A3").setFormula(
    '=IFERROR(SORT({エントリー!B2:B,' +
    ' ARRAYFORMULA(IF(エントリー!C2:C="",,VALUE(LEFT(エントリー!C2:C,FIND("/",エントリー!C2:C)-1)))),' +
    ' ARRAYFORMULA(IF(エントリー!D2:D="",,ABS(VALUE(エントリー!D2:D)-B1))),' +
    ' ARRAYFORMULA(IF(エントリー!E2:E="",,VALUE(REGEXEXTRACT(エントリー!E2:E,"\\((\\d+)\\)"))))},' +
    ' 2,FALSE,3,TRUE,4,TRUE),"(まだエントリーがありません)")'
  );
  r.setFrozenRows(2);
}

function getEntrySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ENTRY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ENTRY_SHEET);
    sh.appendRow(["受信時刻", "名前", "スコア", "ニアピン予想", "完了時刻", "回答パターン", "チェックコード", "検証"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* index.html の makeCode と同一アルゴリズム(djb2) */
function makeCode(name, pattern, guess, finishedMs) {
  var src = name + "|" + pattern + "|" + guess + "|" + finishedMs + "|" + CODE_SALT;
  var h = 5381;
  for (var i = 0; i < src.length; i++) {
    h = ((h * 33) ^ src.charCodeAt(i)) >>> 0;
  }
  return ("00000" + h.toString(36).toUpperCase()).slice(-6);
}

/* 回答パターンから正解数を計算し "正解数/総数" を返す(形式不正ならnull) */
function scoreFromPattern(pattern) {
  var total = 0, correct = 0, ok = true;
  var parts = pattern.split("/");
  if (parts.length !== 3) return null;
  parts.forEach(function (part) {
    var kv = part.split(":");
    var ans = ANSWERS[kv[0]];
    var digits = kv[1] || "";
    if (!ans || digits.length !== ans.length) { ok = false; return; }
    for (var i = 0; i < ans.length; i++) {
      total++;
      if (String(ans[i]) === digits.charAt(i)) correct++;
    }
  });
  return ok ? (correct + "/" + total) : null;
}
