/*************************************************************
 *  Daily Diet Log  →  スプレッドシート自動追記 GAS
 *  ------------------------------------------------------------
 *  使い方（初回のみ・2分）
 *   1. 記録したいスプレッドシートを開く
 *   2. 拡張機能 → Apps Script
 *   3. このコードを全部貼り付けて保存（フロッピー保存）
 *   4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *   5. 「次のユーザーとして実行：自分」「アクセス：全員」→ デプロイ
 *   6. 一度だけ承認 → 表示された「ウェブアプリのURL(.../exec)」をコピー
 *   7. 入力サイトの ⚙設定 に貼る。以上！
 *
 *  ※ 入力サイトの「あいことば」と下の TOKEN は同じ文字にしてください。
 *************************************************************/

var TOKEN = 'merone-diet-2026';        // 入力サイトの token と一致させる
var BUG_SHEET_NAME = '不具合ログ';

// 体組成の並び順（入力サイトと一致）
var BODY_ORDER = [
  { key:'weight', label:'体重(kg)' },
  { key:'fat',    label:'体脂肪(%)' },
  { key:'bmi',    label:'BMI' },
  { key:'water',  label:'水分率(%)' },
  { key:'visc',   label:'内臓脂肪' },
  { key:'sub',    label:'皮下脂肪(%)' },
  { key:'muscle', label:'骨格筋(%)' }
];

/* ============ エントリポイント ============ */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return json({ ok:true, name:'merone-diet', title:ss.getName(), blocks:countBlocks_(ss.getSheets()[0]) });
  } catch (err) {
    return json({ ok:false, error:String(err) });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return json({ ok:false, error:'あいことば(token)が違います' });

    if (data.action === 'bug') { logBug_(data); return json({ ok:true }); }

    // 既定：1日ぶんを保存
    var row = saveDay_(data);
    return json({ ok:true, row:row, message:'保存しました' });
  } catch (err) {
    return json({ ok:false, error:String(err) });
  }
}

/* ============ 1日ブロックを追記 ============ */
function saveDay_(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var blockNo = countBlocks_(sheet) + 1;

  var meals  = d.meals || [];
  var care   = d.care  || {};
  var I      = normIdeal_(d.ideal);
  var body   = d.body  || {};

  // 合計
  var sum = { kcal:0, pro:0, water:0 };
  meals.forEach(function(m){ sum.kcal += n_(m.kcal); sum.pro += n_(m.protein); sum.water += n_(m.water); });
  var sleepH = n_(care.sleepH), bath = n_(care.bath), aero = n_(care.aero), mas = n_(care.mas);

  // 前日・当初（既存ブロックから）
  var refs = getBodyRefs_(sheet);   // {prev:{}, base:{}}

  var R = [];  // 9列 × N行
  R.push(row9_('Day' + blockNo, fmtDate_(d.date), '', '', '', '', '', '氏名', d.name || ''));
  R.push(blank9_());
  R.push(row9_('', '', '', '', '', '', '', '排便有無', d.haiben || ''));
  R.push(row9_('', '', '', '', '', '', '', '便秘薬摂取', d.benpi || ''));
  R.push(row9_('', '', '', '', '', '', '', '体調', d.taicho || ''));
  R.push(row9_('時刻','食事','カロリー(kcal)','プロテイン(g)','水(ml)','睡眠時間','入浴(m)','有酸素運動(m)','マッサージ(m)'));

  // 起床行：入浴・有酸素・マッサージの合計をここに置く（列合計が総量と一致）
  R.push(row9_(fmtTime_(d.wake), '', '', '', '', '起床', bath || '', aero || '', mas || ''));

  // 食事の行
  meals.forEach(function(m){
    R.push(row9_(fmtTime_(m.time), m.food || '', blankNum_(m.kcal), blankNum_(m.protein), blankNum_(m.water), '', '', '', ''));
  });

  // 就寝行
  R.push(row9_(fmtTime_(d.sleepTime), '', '', '', '', '就寝', '', '', ''));

  // 集計ブロック
  R.push(row9_('→食事回数：', meals.length, '', '', '', (sleepH ? sleepH + '時間' : ''), '', '', ''));
  R.push(row9_('', '合計', sum.kcal, round1_(sum.pro), sum.water, sleepH, bath, aero, mas));
  R.push(row9_('', '理想', I.kcal, I.pro, I.water, I.sleep, I.bath, I.aero, I.mas));
  R.push(row9_('', '差分', sum.kcal - I.kcal, round1_(sum.pro - I.pro), sum.water - I.water, sleepH - I.sleep, bath - I.bath, aero - I.aero, mas - I.mas));
  R.push(row9_('', '累計', sum.kcal, round1_(sum.pro), sum.water, sleepH, bath, '', ''));
  R.push(blank9_());

  // 体組成
  R.push(row9_('', '前日', '当日', '前日比', '当初比*', '目標', '差分(残り)', '', ''));
  BODY_ORDER.forEach(function(b){
    var cur  = pick_(body[b.key], 'cur');
    var goal = pick_(body[b.key], 'goal');
    var prev = refs.prev[b.key];
    var base = refs.base[b.key];
    if (base === undefined || base === null) base = cur;   // 初回は当日＝当初
    var zenpi   = (prev != null && cur != null) ? round2_(cur - prev) : '';
    var nokori  = (goal != null && cur != null) ? round2_(goal - cur) : '';
    R.push(row9_(b.label, nz_(prev), nz_(cur), zenpi, nz_(base), nz_(goal), nokori, '', ''));
  });
  R.push(row9_('', '', '', '', '*Day 0との比較', '', '', '', ''));

  // 追記位置（既存の下に2行あけて）
  var last = sheet.getLastRow();
  var start = (last > 0) ? last + 3 : 1;
  sheet.getRange(start, 1, R.length, 9).setValues(R);

  formatBlock_(sheet, start, R.length);
  return start;
}

/* ============ 既存ブロック参照（前日・当初） ============ */
function getBodyRefs_(sheet) {
  var out = { prev:{}, base:{} };
  var last = sheet.getLastRow();
  if (last < 1) return out;
  var colA = sheet.getRange(1, 1, last, 1).getValues();       // ラベル列
  var colC = sheet.getRange(1, 3, last, 1).getValues();       // 当日列
  var weightRows = [];
  for (var i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).indexOf('体重') === 0) weightRows.push(i);   // 0-based
  }
  if (!weightRows.length) return out;
  var firstW = weightRows[0];
  var lastW  = weightRows[weightRows.length - 1];
  BODY_ORDER.forEach(function(b, idx){
    var pv = colC[lastW + idx] ? colC[lastW + idx][0] : '';
    var bv = colC[firstW + idx] ? colC[firstW + idx][0] : '';
    out.prev[b.key] = (pv === '' ? null : Number(pv));
    out.base[b.key] = (bv === '' ? null : Number(bv));
  });
  return out;
}

function countBlocks_(sheet) {
  var last = sheet.getLastRow();
  if (last < 1) return 0;
  var colH = sheet.getRange(1, 8, last, 1).getValues();   // 「氏名」がブロックに1つ
  var c = 0;
  for (var i = 0; i < colH.length; i++) if (String(colH[i][0]).indexOf('氏名') === 0) c++;
  return c;
}

/* ============ 見た目を整える ============ */
function formatBlock_(sheet, start, len) {
  try {
    // ブロック全体うっすら枠
    sheet.getRange(start, 1, len, 9).setVerticalAlignment('middle');
    // タイトル行
    sheet.getRange(start, 1, 1, 9).setBackground('#f7c6d4').setFontWeight('bold');
    // 表ヘッダー（時刻…の行 = start+5）
    sheet.getRange(start + 5, 1, 1, 9).setBackground('#fdeef2').setFontWeight('bold').setHorizontalAlignment('center');
    // ラベル系太字
    sheet.getRange(start + 2, 8, 3, 1).setFontWeight('bold');   // 排便/便秘薬/体調
    // 集計ラベル列を太字（合計/理想/差分/累計 が入る B 列付近）
  } catch (e) { /* 書式は失敗しても致命的でない */ }
}

/* ============ 不具合ログ ============ */
function logBug_(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BUG_SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(BUG_SHEET_NAME); sh.appendRow(['日時','内容','端末','ページ']); }
  sh.appendRow([new Date(), d.text || '', d.ua || '', d.page || '']);
}

/* ============ ユーティリティ ============ */
function json(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function n_(v){ v = parseFloat(v); return isNaN(v) ? 0 : v; }
function nz_(v){ return (v === null || v === undefined || v === '') ? '' : v; }
function blankNum_(v){ v = parseFloat(v); return isNaN(v) ? '' : v; }
function pick_(o, k){ return (o && o[k] !== undefined && o[k] !== null && o[k] !== '') ? Number(o[k]) : null; }
function round1_(v){ return Math.round(n_(v) * 10) / 10; }
function round2_(v){ return Math.round(n_(v) * 100) / 100; }
function row9_(a,b,c,d,e,f,g,h,i){ return [a,b,c,d,e,f,g,h,i]; }
function blank9_(){ return ['','','','','','','','','']; }
function normIdeal_(I){
  I = I || {};
  return { kcal:n_(I.kcal)||470, pro:n_(I.pro)||80, water:n_(I.water)||3000,
           sleep:n_(I.sleep)||7, bath:n_(I.bath)||60, aero:n_(I.aero)||60, mas:n_(I.mas)||60 };
}
function fmtDate_(iso){
  if (!iso) return '';
  var p = String(iso).split('-');
  if (p.length !== 3) return iso;
  return Number(p[0]) + '.' + Number(p[1]) + '.' + Number(p[2]);   // 2026.8.29
}
function fmtTime_(t){
  if (!t) return '';
  var p = String(t).split(':');
  if (p.length < 2) return t;
  return Number(p[0]) + ':' + p[1];   // 07:00 -> 7:00
}

/* ============ 動作テスト（任意）============ */
function testWrite() {
  var demo = {
    token: TOKEN, action:'saveDay', date:'2026-08-29', name:'テスト',
    haiben:'あり', benpi:'なし', taicho:'スッキリ', wake:'07:00', sleepTime:'00:00',
    care:{ sleepH:7, bath:60, aero:30, mas:60 },
    ideal:{ kcal:470, pro:80, water:3000, sleep:7, bath:60, aero:60, mas:60 },
    meals:[ {time:'08:00',food:'ツナボール',kcal:47,protein:8,water:200},
            {time:'12:00',food:'鶏胸',kcal:120,protein:25,water:300} ],
    body:{ weight:{cur:60.2,goal:47}, fat:{cur:31.8,goal:22}, bmi:{cur:22.38,goal:23},
           water:{cur:46.7,goal:57}, visc:{cur:5,goal:7}, sub:{cur:29.5,goal:19}, muscle:{cur:39.7,goal:null} }
  };
  Logger.log('書込行: ' + saveDay_(demo));
}
