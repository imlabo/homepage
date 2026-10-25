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

function doGet() {
  return out_(readAll_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const pass = PropertiesService.getScriptProperties().getProperty("PASSCODE");
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
