/**
 * 雀卓ランキング：保存先スプレッドシート用 Apps Script
 *
 * 使い方
 * 1. 新しい Google スプレッドシートを作る
 * 2. 拡張機能 > Apps Script を開き、このファイルの中身を全部貼り付けて保存
 * 3. （任意）プロジェクトの設定 > スクリプト プロパティ に PASSCODE を追加すると、
 *    記録・削除のときに合言葉が必要になる（閲覧は誰でも可）
 * 4. デプロイ > 新しいデプロイ > 種類「ウェブアプリ」
 *    次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * 5. 表示された「ウェブアプリのURL」を mahjong/config.js の apiUrl に貼る
 *
 * 写真からの自動読み取り（任意）
 * - Google AI Studio（https://aistudio.google.com/apikey）で API キーを作り、
 *   スクリプト プロパティ GEMINI_API_KEY に設定する（無料枠あり）
 * - モデルを変えたいときは GEMINI_MODEL（既定 gemini-2.5-flash）
 * - コードを更新したら「デプロイを管理 → 編集 → バージョン：新バージョン → デプロイ」
 *   （URLを変えずに反映できる。「新しいデプロイ」だとURLが変わるので注意）
 *
 * シート games / yakuman / config は初回アクセス時に自動で作られる。
 * 最後の json 列がアプリの読み込み元。ほかの列は人が見るための写し。
 */

const SHEETS = {
  games: ["id", "登録日時", "対局日",
          "東家", "東家 素点", "東家 pt", "南家", "南家 素点", "南家 pt",
          "西家", "西家 素点", "西家 pt", "北家", "北家 素点", "北家 pt", "json"],
  yakuman: ["id", "登録日時", "日付", "あがった人", "役満", "ツモ/ロン", "放銃者", "メモ", "json"],
  config: ["key", "json"],
};

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(SHEETS[name]);
    s.setFrozenRows(1);
  }
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readAll_() {
  const res = { games: [], yakuman: [], config: {} };
  ["games", "yakuman"].forEach(function (n) {
    const v = sheet_(n).getDataRange().getValues();
    const jc = SHEETS[n].length - 1;
    for (let i = 1; i < v.length; i++) {
      try {
        const o = JSON.parse(v[i][jc]);
        o.id = String(v[i][0]);
        res[n].push(o);
      } catch (e) {}
    }
  });
  const c = sheet_("config").getDataRange().getValues();
  for (let i = 1; i < c.length; i++) {
    try { res.config[c[i][0]] = JSON.parse(c[i][1]); } catch (e) {}
  }
  return res;
}

function doGet(e) {
  // 接続確認：ブラウザで「ウェブアプリのURL?test=gemini」を開く
  if (e && e.parameter && e.parameter.test === "gemini") {
    const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!key) return out_({ error: "no_api_key", hint: "スクリプト プロパティに GEMINI_API_KEY がありません" });
    try {
      const g = gemini_(key, [{ text: "「OK」とだけ答えて" }], false);
      if (!g.ok) return out_(g);
      return out_({ ok: true, model: g.model, reply: g.res.getContentText().slice(0, 200) });
    } catch (err) {
      return out_({ error: "exception", detail: String(err) });
    }
  }
  return out_(readAll_());
}

/* ---- 写真から点数を読み取る（Gemini API・無料枠で利用可） ---- */
const READ_PROMPT = [
  "これは全自動麻雀卓の点数表示（液晶）の写真です。4人の持ち点を読み取ってください。",
  "表示の配置:",
  "- 中央下の一番大きい数字 = 手前（撮影者）の持ち点 → bottom",
  "- 右側の大きい数字 → right、上側中央の大きい数字 → top、左側の大きい数字 → left",
  "- 右・上・左の大きい数字の下にある小さい数字（記号付き）は手前との点差。記号は無視して絶対値を diffRight diffTop diffLeft に。無ければ null",
  "- 右上の CHECK の数字は合計 → check（無ければ null）",
  "- 1ST/2ND/3RD/4TH などの順位表示は無視",
  "持ち点がマイナスの人は負の数に。7セグメント表示なので 0と8、1と7、5と6 の読み違いに注意し、点差・合計と矛盾しないか確かめること。",
  'JSONのみで答える。例: {"bottom":31500,"right":24600,"top":27100,"left":16800,"diffRight":6900,"diffTop":4400,"diffLeft":14700,"check":100000}',
  '点数表示が写っていない・読めない場合は {"error":"理由"}'
].join("\n");

function ver_(n) {
  const m = n.match(/gemini-(\d+(?:\.\d+)?)/);
  return m ? parseFloat(m[1]) : 0;
}

/* 今使える Flash 系モデルを Google に問い合わせ、新しい順に並べる（6時間キャッシュ） */
function models_(key) {
  const fixed = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
  const cache = CacheService.getScriptCache();
  let listed = null;
  const c = cache.get("gemini_models");
  if (c) listed = JSON.parse(c);
  if (!listed) {
    listed = [];
    const r = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models?pageSize=200",
      { headers: { "x-goog-api-key": key }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      const all = (JSON.parse(r.getContentText()).models || [])
        .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf("generateContent") >= 0; })
        .map(function (m) { return m.name.replace(/^models\//, ""); })
        .filter(function (n) { return /^gemini-.*flash/.test(n) && !/(lite|tts|image|live|audio|embed|exp)/.test(n); });
      const stable = all.filter(function (n) { return !/preview/.test(n); });
      listed = (stable.length ? stable : all).sort(function (a, b) { return ver_(b) - ver_(a) || a.length - b.length; }).slice(0, 3);
      if (listed.length) cache.put("gemini_models", JSON.stringify(listed), 21600);
    }
  }
  const list = [fixed].concat(listed, ["gemini-flash-latest"]).filter(Boolean);
  return list.filter(function (x, i) { return list.indexOf(x) === i; });
}

/* モデルが廃止・混雑していたら次の候補で再試行する */
function gemini_(key, parts, json) {
  let last = null;
  const errors = [];
  const models = models_(key);
  for (let i = 0; i < models.length; i++) {
    const res = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/" + models[i] + ":generateContent",
      {
        method: "post",
        contentType: "application/json",
        headers: { "x-goog-api-key": key },
        muteHttpExceptions: true,
        payload: JSON.stringify({
          contents: [{ parts: parts }],
          generationConfig: json ? { responseMimeType: "application/json", temperature: 0 } : { temperature: 0 }
        })
      }
    );
    const code = res.getResponseCode();
    if (code === 200) return { ok: true, model: models[i], res: res };
    errors.push(models[i] + ":" + code);
    last = { error: "api_error", model: models[i], status: code, detail: "試したモデル " + errors.join(", ") + " / " + res.getContentText().slice(0, 240) };
    if (code === 400 || code === 401 || code === 403) break; // キーの問題は他のモデルでも同じ
  }
  return last;
}

function readScores_(req) {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key) return { error: "no_api_key" };
  if (!req.image || String(req.image).length > 8000000) return { error: "bad_image" };
  const g = gemini_(key, [
    { inline_data: { mime_type: req.mime || "image/jpeg", data: req.image } },
    { text: READ_PROMPT }
  ], true);
  if (!g.ok) return g;
  const res = g.res;
  try {
    const body = JSON.parse(res.getContentText());
    const text = body.candidates[0].content.parts.map(function (p) { return p.text || ""; }).join("");
    const m = text.match(/\{[\s\S]*\}/);
    const j = JSON.parse(m ? m[0] : text);
    if (j.error) return { error: "unreadable", detail: String(j.error) };
    return { ok: true, scores: j };
  } catch (err) {
    return { error: "parse_error" };
  }
}

function doPost(e) {
  const req0 = JSON.parse(e.postData.contents);
  const pass0 = PropertiesService.getScriptProperties().getProperty("PASSCODE");
  if (req0.action === "read") {
    if (pass0 && String(req0.pass || "") !== pass0) return out_({ error: "passcode" });
    try {
      return out_(readScores_(req0)); // 読み取りは書き込みロック不要
    } catch (err) {
      return out_({ error: "exception", detail: String(err) });
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = req0;
    const pass = pass0;
    if (pass && String(req.pass || "") !== pass) return out_({ error: "passcode" });

    if (req.action === "add") {
      const n = req.coll;
      const it = req.item || {};
      const now = new Date();
      let row;
      if (n === "games") {
        if (!Array.isArray(it.players) || it.players.length !== 4) return out_({ error: "bad_item" });
        const id = "g" + Utilities.getUuid().slice(0, 8);
        row = [id, now, it.date];
        it.players.slice().sort(function (a, b) { return a.seat - b.seat; })
          .forEach(function (p) { row.push(p.name, p.score, p.pt); });
        row.push(JSON.stringify(it));
        sheet_(n).appendRow(row);
        return out_({ ok: true, id: id });
      }
      if (n === "yakuman") {
        if (!it.who || !Array.isArray(it.yaku)) return out_({ error: "bad_item" });
        const id = "y" + Utilities.getUuid().slice(0, 8);
        row = [id, now, it.date, it.who, it.yaku.join("・"), it.how, it.from || "", it.memo || "", JSON.stringify(it)];
        sheet_(n).appendRow(row);
        return out_({ ok: true, id: id });
      }
      return out_({ error: "bad_collection" });
    }

    if (req.action === "del") {
      if (req.coll !== "games" && req.coll !== "yakuman") return out_({ error: "bad_collection" });
      const s = sheet_(req.coll);
      const last = s.getLastRow();
      if (last > 1) {
        const ids = s.getRange(2, 1, last - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === String(req.id)) s.deleteRow(i + 2);
        }
      }
      return out_({ ok: true });
    }

    if (req.action === "setConfig") {
      if (req.key !== "rule" && req.key !== "table") return out_({ error: "bad_key" });
      const s = sheet_("config");
      const v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (v[i][0] === req.key) {
          s.getRange(i + 1, 2).setValue(JSON.stringify(req.value));
          return out_({ ok: true });
        }
      }
      s.appendRow([req.key, JSON.stringify(req.value)]);
      return out_({ ok: true });
    }

    return out_({ error: "unknown_action" });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 初回だけエディタで実行する：外部通信（Gemini）の許可を出すための関数。
 * 関数の選択欄で authorize を選び「実行」→ 承認画面で「許可」。
 */
function authorize() {
  const r = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/", { muteHttpExceptions: true });
  Logger.log("外部通信OK: HTTP " + r.getResponseCode());
}
