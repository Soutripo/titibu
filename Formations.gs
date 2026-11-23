/**
 * 編成のサーバー同期(サーバー側)。Code.gs / Extras.gs とは別のファイル(Formations.gs)として追加してください。
 *
 * ・運行図表の「運用表」で入れた編成(行路番号 → 編成文字列)を、日付ごとにサーバーへ保存します。
 * ・見る側(全員): ログインできていれば、その日の編成を取得できます(窓口 fget)。
 * ・入れる側(管理者だけ): 編成の保存ができます(窓口 fput)。Extras.gs の xput と同じ考え方です。
 * ・保存先は「編成」シート。1日 = 1行で、中身は JSON のまま持っています(臨時列車シートと同じ方式)。
 *
 * 入れ方: README の「編成のサーバー同期」を見てください。
 *   1. このファイルを Formations.gs として追加する
 *   2. Extras.gs の doPost の switch に、fget / fput の2行を足す
 *   3. setupFormations を一度だけ実行する
 *   4. ウェブアプリを新しいバージョンでデプロイする
 */
const SHEET_FORMS = '編成';
const FM_KEEP_DAYS = 60;       // 日付からこの日数を過ぎた記録は、保存操作のときに自動で消します(臨時列車と同じ)
const FM_AHEAD_DAYS = 31;      // これより先の日付の編成は登録できません
const FM_PAST_DAYS = 7;        // これより前の日付の編成は登録できません
const FM_HEAD = ['日付', 'ダイヤ', '内容(消さないでください)', '登録者', '登録日時'];

/* ===== 初回設定: 一度だけ実行してください(何度実行しても安全です) ===== */
function setupFormations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_FORMS);
  if (!sh) sh = ss.insertSheet(SHEET_FORMS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, FM_HEAD.length).setValues([FM_HEAD]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');   // 日付の列は、日付に変換されない文字列にする
  try {
    SpreadsheetApp.getUi().alert('編成の設定が完了しました。\n編成を登録できるのは「' + SHEET_USERS + '」シートの「管理者」列に ○ が入っているIDだけです。');
  } catch (e) {
    console.log('編成の設定が完了しました。');
  }
}

/* ===== 保存シートの読み書き ===== */
function formsSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FORMS); }
function readForms_() {
  const sh = formsSheet_();
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, FM_HEAD.length).getValues(), out = [];
  vals.forEach(function (v, i) {
    try {
      const o = JSON.parse(String(v[2]));
      if (o && /^\d{4}-\d{2}-\d{2}$/.test(String(v[0])) && (o.day === 'weekday' || o.day === 'holiday') && o.forms && typeof o.forms === 'object') {
        out.push({ date: String(v[0]), day: o.day, forms: o.forms, _row: i + 2, _by: String(v[3]), _at: v[4] instanceof Date ? v[4].getTime() : 0 });
      }
    } catch (_) {}     // 手で壊れた行は読み飛ばす
  });
  return out;
}
function purgeOldForms_(sh) {
  const limit = jstDate_(-FM_KEEP_DAYS);
  readForms_().filter(function (r) { return r.date < limit; }).sort(function (p, q) { return q._row - p._row; }).forEach(function (r) { sh.deleteRow(r._row); });
}

/* 見る側: 今日(日本時間)の編成を返す。23:55 以降は、端末の時計が少し進んでいる場合のため翌日分の要求も受け付ける(臨時列車と同じ) */
function formationsGet_(req) {
  const a = auth_(req, false);
  if (a.err) return failOf_(a.err);
  const today = jstDate_(0), want = String(req.d || '');
  const d = (want === jstDate_(1) && jstMinutes_() >= 23 * 60 + 55) ? want : today;
  const f = readForms_().filter(function (r) { return r.date === d; })[0];
  return { ok: true, date: d, tok: makeToken_(a.u.id, String(req.dev)), day: f ? f.day : null, forms: f ? f.forms : {}, by: f ? f._by : '', at: f ? f._at : 0 };
}

/* 入れる側(管理者だけ): その日の編成を丸ごと保存する */
function formationsAdmin_(req) {
  const a = auth_(req, true);
  if (a.err) return failOf_(a.err);
  const sh = formsSheet_();
  if (!sh) return { ok: false, msg: '「' + SHEET_FORMS + '」シートがありません。setupFormations を実行してください' };
  const c = cleanForms_(req);
  if (c.error) return { ok: false, msg: c.error };
  purgeOldForms_(sh);
  const res = { ok: true, tok: makeToken_(a.u.id, String(req.dev)), date: c.date };
  const f = readForms_().filter(function (r) { return r.date === c.date; })[0];
  const row = [c.date, c.day === 'weekday' ? '平日' : '休日', JSON.stringify({ day: c.day, forms: c.forms }), a.u.id, new Date()];
  if (f) { sh.getRange(f._row, 1).setNumberFormat('@'); sh.getRange(f._row, 1, 1, FM_HEAD.length).setValues([row]); }
  else { sh.appendRow(row); sh.getRange(sh.getLastRow(), 1).setNumberFormat('@'); }
  log_(a.u.id, '編成を' + (f ? '更新' : '登録') + ': ' + c.date + ' ' + Object.keys(c.forms).length + '行路', '');
  return res;
}

/* 登録内容の確認(日付・ダイヤ・編成の形。行路番号が実在するかは、読み込む側の運行図表で確認する) */
function cleanForms_(req) {
  const date = String(req.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || Utilities.formatDate(new Date(date + 'T12:00:00+09:00'), 'Asia/Tokyo', 'yyyy-MM-dd') !== date) return { error: '日付が正しくありません' };
  if (date < jstDate_(-FM_PAST_DAYS) || date > jstDate_(FM_AHEAD_DAYS)) return { error: 'その日付の編成は登録できません' };
  const day = String(req.day || '');
  if (day !== 'weekday' && day !== 'holiday') return { error: 'ダイヤ(平日・休日)が正しくありません' };
  const src = req.forms;
  if (!src || typeof src !== 'object' || Array.isArray(src)) return { error: '編成のデータがありません' };
  const keys = Object.keys(src);
  if (keys.length > 200) return { error: '行路が多すぎます' };
  const forms = {};
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (!/^[A-Za-z0-9_-]{1,12}$/.test(k)) return { error: '行路番号が正しくありません: ' + k };
    const v = String(src[k] == null ? '' : src[k]).replace(/[\u0000-\u001f]/g, ' ').trim();
    if (!v) continue;                       // 空の編成は保存しない(消す操作として扱う)
    if (v.length > 30) return { error: '編成が長すぎます(30文字まで): ' + k };
    forms[k] = v;
  }
  return { date: date, day: day, forms: forms };
}
