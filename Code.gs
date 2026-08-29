/*************************************************************
 *  Daily Diet Log v2  →  1日1タブに「1件ずつ」追記するGAS
 *  ------------------------------------------------------------
 *  貼り方（初回/更新とも同じ）
 *   1. 記録スプレッドシートを開く → 拡張機能 → Apps Script
 *   2. このコードを全部貼り付けて保存
 *   3. デプロイ → デプロイを管理 → 鉛筆✎ → バージョン「新バージョン」→ デプロイ
 *      （初回は「新しいデプロイ」→ ウェブアプリ / 実行:自分 / アクセス:全員）
 *   ※ URLは変わりません。あいことば(TOKEN)は下と入力サイトを一致させる
 *************************************************************/

var TOKEN = 'merone-diet-2026';
var TZ = 'Asia/Tokyo';
var HEADER = ['時刻','食べたもの','カロリー(kcal)','プロテイン(g)','水分(ml)','体重(kg)','睡眠(h)','メモ'];
var BUG_SHEET = '不具合ログ';

/* ===== entry points ===== */
function doGet(e){
  try{
    var p=(e&&e.parameter)||{};
    if(p.date) return json(getDay_(p.date));
    return json({ok:true, name:'merone-diet2', title:ss_().getName()});
  }catch(err){ return json({ok:false, error:String(err)}); }
}

function doPost(e){
  try{
    var d=JSON.parse(e.postData.contents);
    if(d.token!==TOKEN) return json({ok:false, error:'あいことば(token)が違います'});
    switch(d.action){
      case 'addEntry':    return json(addEntry_(d));
      case 'getDay':      return json(getDay_(d.date));
      case 'deleteEntry': return json(deleteEntry_(d));
      case 'setupTabs':   return json(setupTabs_());
      case 'bug':         logBug_(d); return json({ok:true});
      default:            return json({ok:false, error:'unknown action: '+d.action});
    }
  }catch(err){ return json({ok:false, error:String(err)}); }
}

/* ===== 期の判定・タブ名 ===== */
function phaseOf_(date){
  if(date>='2026-08-28' && date<='2026-08-30') return '分析期';
  if(date>='2026-08-31' && date<='2026-09-10') return '減量期';
  if(date>='2026-09-11' && date<='2026-09-27') return '分析期';
  return '';
}
function tabName_(date){
  var p=String(date).split('-');
  var name=Number(p[1])+'/'+Number(p[2]);
  var ph=phaseOf_(date);
  return ph ? name+' '+ph : name;
}
function findTab_(date){
  var ss=ss_(), name=tabName_(date);
  return ss.getSheetByName(name) || ss.getSheetByName(name.replace(/\//g,'-'));
}
function getOrCreateDayTab_(date){
  var sh=findTab_(date);
  if(sh) return sh;
  var ss=ss_(), name=tabName_(date), created;
  try{ created=ss.insertSheet(name); }
  catch(e){ name=name.replace(/\//g,'-'); created=ss.getSheetByName(name)||ss.insertSheet(name); }
  initTab_(created, date);
  return created;
}
function initTab_(sh, date){
  sh.getRange(1,1,1,HEADER.length).setValues([HEADER]).setFontWeight('bold').setBackground('#f7c6d4').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  try{ sh.setColumnWidth(1,64); sh.setColumnWidth(2,200); sh.setColumnWidth(8,180); }catch(e){}
  try{ sh.getRange('A1').setNote('日付: '+date+' / '+(phaseOf_(date)||'期間外')); }catch(e){}
}

/* ===== 1件追記 ===== */
function addEntry_(d){
  var sh=getOrCreateDayTab_(d.date);
  sh.appendRow([ d.time||'', d.food||'', bn_(d.kcal), bn_(d.protein), bn_(d.water), bn_(d.weight), bn_(d.sleep), d.memo||'' ]);
  return objAssign_({ok:true}, getDay_(d.date));
}

/* ===== その日の一覧＋合計 ===== */
function getDay_(date){
  var sh=findTab_(date);
  var out={ ok:true, date:date, tab:tabName_(date), rows:[], total:{kcal:0,protein:0,water:0}, count:0, weight:'' };
  if(!sh) return out;
  var last=sh.getLastRow();
  if(last<2) return out;
  var vals=sh.getRange(2,1,last-1,HEADER.length).getValues();
  for(var i=0;i<vals.length;i++){
    var r=vals[i];
    if(String(r.join(''))==='') continue;
    out.rows.push({ row:i+2, time:fmtCell_(r[0]), food:r[1], kcal:n_(r[2]), protein:n_(r[3]), water:n_(r[4]), weight:r[5]===''?'':n_(r[5]), sleep:r[6]===''?'':n_(r[6]), memo:r[7] });
    out.total.kcal+=n_(r[2]); out.total.protein+=n_(r[3]); out.total.water+=n_(r[4]);
    if(r[5]!=='' && out.weight==='') out.weight=n_(r[5]);
  }
  out.total.protein=Math.round(out.total.protein*10)/10;
  out.count=out.rows.length;
  return out;
}

/* ===== 1件削除 ===== */
function deleteEntry_(d){
  var sh=findTab_(d.date);
  if(sh && d.row>=2 && d.row<=sh.getLastRow()) sh.deleteRow(d.row);
  return objAssign_({ok:true}, getDay_(d.date));
}

/* ===== 8/28〜9/27のタブを一括生成 ===== */
function setupTabs_(){
  var made=[];
  for(var i=0;i<31;i++){
    var dt=new Date(2026,7,28+i,12,0,0);                 // Aug=7. 8/28 + i
    var date=Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    if(!findTab_(date)){ getOrCreateDayTab_(date); made.push(tabName_(date)); }
  }
  return { ok:true, created:made.length, tabs:made };
}

/* ===== bug ===== */
function logBug_(d){
  var ss=ss_(), sh=ss.getSheetByName(BUG_SHEET);
  if(!sh){ sh=ss.insertSheet(BUG_SHEET); sh.appendRow(['日時','内容','端末','ページ']); }
  sh.appendRow([new Date(), d.text||'', d.ua||'', d.page||'']);
}

/* ===== utils ===== */
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function n_(v){ v=parseFloat(v); return isNaN(v)?0:v; }
function bn_(v){ v=parseFloat(v); return isNaN(v)?'':v; }
function fmtCell_(v){
  if(v instanceof Date){ return Utilities.formatDate(v, TZ, 'H:mm'); }
  return v;
}
function objAssign_(a,b){ for(var k in b){ if(b.hasOwnProperty(k)) a[k]=b[k]; } return a; }

/* ===== 動作テスト（任意）===== */
function testAll(){
  Logger.log(setupTabs_());
  Logger.log(addEntry_({date:'2026-08-29', time:'8:00', food:'テスト', kcal:100, protein:5, water:200, weight:60.2, sleep:7, memo:'テスト'}));
  Logger.log(getDay_('2026-08-29'));
}
