/*************************************************************
 *  Daily Diet Log v3  →  各日タブ=「分析ブロック(上)＋食事ログ(下)」
 *  ------------------------------------------------------------
 *  貼り方（更新も同じ）
 *   1. スプレッドシート → 拡張機能 → Apps Script
 *   2. 全部貼り付けて保存
 *   3. デプロイ → デプロイを管理 → 鉛筆✎ → バージョン「新バージョン」→ デプロイ
 *      （URLは変わりません。TOKENは入力サイトと一致させる）
 *
 *  タブ構成（各日）:
 *   1行目 : タイトル（例 8月28日（金） 分析期）
 *   3-10行: 体組成テーブル（前日・当日・前日比・当初比・目標・差分）
 *   12-16行: カロリー等の 合計/理想/差分/累計（合計はログから自動集計）
 *   18行目 : ログ見出し（時刻/食べたもの/カロリー/プロテイン/水分/メモ）
 *   19行目〜: 記録（1件=1行・下に追記）
 *************************************************************/

var TOKEN = 'merone-diet-2026';
var TZ = 'Asia/Tokyo';
var BUG_SHEET = '不具合ログ';

var LOG_HEADER = ['時刻','食べたもの','カロリー(kcal)','プロテイン(g)','水分(ml)','メモ'];
var LOG_HEAD_ROW = 18;
var LOG_START = 19;

// 体組成（key,label,目標）  行4〜10
var BODY = [
  ['weight','体重(kg)',47],
  ['fat','体脂肪(%)',22],
  ['bmi','BMI',22],
  ['water','水分率(%)',57],
  ['visc','内臓脂肪',7],
  ['sub','皮下脂肪(%)',19],
  ['muscle','骨格筋(%)','']
];
var BODY_START = 4;
var SLEEP_CELL = [13,5];   // 合計行の睡眠(h)  E13

/* ===== entry points ===== */
function doGet(e){
  try{
    var p=(e&&e.parameter)||{};
    if(p.date) return json(getDay_(p.date));
    return json({ok:true, name:'merone-diet3', title:ss_().getName()});
  }catch(err){ return json({ok:false, error:String(err)}); }
}
function doPost(e){
  try{
    var d=JSON.parse(e.postData.contents);
    if(d.token!==TOKEN) return json({ok:false, error:'あいことば(token)が違います'});
    switch(d.action){
      case 'addEntry':    return json(addEntry_(d));
      case 'saveBody':    return json(saveBody_(d));
      case 'getDay':      return json(getDay_(d.date));
      case 'deleteEntry': return json(deleteEntry_(d));
      case 'setupTabs':   return json(setupTabs_());
      case 'bug':         logBug_(d); return json({ok:true});
      default:            return json({ok:false, error:'unknown action: '+d.action});
    }
  }catch(err){ return json({ok:false, error:String(err)}); }
}

/* ===== 期・タブ名 ===== */
function phaseOf_(date){
  if(date>='2026-08-28' && date<='2026-08-30') return '分析期';
  if(date>='2026-08-31' && date<='2026-09-10') return '減量期';
  if(date>='2026-09-11' && date<='2026-09-27') return '分析期';
  return '';
}
function tabName_(date){
  var p=String(date).split('-');
  var nm=Number(p[1])+'/'+Number(p[2]);
  var ph=phaseOf_(date);
  return ph ? nm+' '+ph : nm;
}
function tabTitle_(date){
  var p=String(date).split('-');
  var d=new Date(Number(p[0]),Number(p[1])-1,Number(p[2]));
  var dow=['日','月','火','水','木','金','土'][d.getDay()];
  var ph=phaseOf_(date);
  return Number(p[1])+'月'+Number(p[2])+'日（'+dow+'）'+(ph?'　'+ph:'');
}
function findTab_(date){
  var ss=ss_(), name=tabName_(date);
  return ss.getSheetByName(name) || ss.getSheetByName(name.replace(/\//g,'-'));
}
function getOrCreateDayTab_(date){
  var sh=findTab_(date);
  if(sh) return sh;
  var ss=ss_(), name=tabName_(date), c;
  try{ c=ss.insertSheet(name); }
  catch(e){ name=name.replace(/\//g,'-'); c=ss.getSheetByName(name)||ss.insertSheet(name); }
  buildTemplate_(c, date);
  return c;
}

/* ===== テンプレ生成（分析ブロック） ===== */
function buildTemplate_(sh, date){
  sh.clear();
  // タイトル
  sh.getRange(1,1,1,7).merge().setValue(tabTitle_(date)).setFontWeight('bold').setBackground('#f7c6d4').setHorizontalAlignment('center');
  // 体組成ヘッダー
  sh.getRange(3,1,1,7).setValues([['からだ（朝の計測）','前日','当日','前日比','当初比*','目標','差分(残り)']])
    .setFontWeight('bold').setBackground('#fdeef2').setHorizontalAlignment('center');
  sh.getRange(3,1).setHorizontalAlignment('left');
  // 体組成 行
  for(var i=0;i<BODY.length;i++){
    var r=BODY_START+i;
    sh.getRange(r,1).setValue(BODY[i][1]).setFontWeight('bold');
    if(BODY[i][2]!=='') sh.getRange(r,6).setValue(BODY[i][2]);   // 目標=F列
  }
  // 合計ブロック ヘッダー
  sh.getRange(12,1,1,5).setValues([['1日の合計','カロリー(kcal)','プロテイン(g)','水分(ml)','睡眠(h)']])
    .setFontWeight('bold').setBackground('#fdeef2').setHorizontalAlignment('center');
  sh.getRange(12,1).setHorizontalAlignment('left');
  sh.getRange(13,1).setValue('合計').setFontWeight('bold');
  sh.getRange(14,1).setValue('理想').setFontWeight('bold');
  sh.getRange(15,1).setValue('差分').setFontWeight('bold');
  sh.getRange(16,1).setValue('累計').setFontWeight('bold');
  // 合計＝ログ自動集計（C/D/E列を19行目以降SUM）
  sh.getRange(13,2).setFormula('=SUM(C'+LOG_START+':C)');
  sh.getRange(13,3).setFormula('=SUM(D'+LOG_START+':D)');
  sh.getRange(13,4).setFormula('=SUM(E'+LOG_START+':E)');
  // 理想
  sh.getRange(14,2,1,4).setValues([[470,80,3000,7]]);
  // 差分
  sh.getRange(15,2).setFormula('=B13-B14');
  sh.getRange(15,3).setFormula('=C13-C14');
  sh.getRange(15,4).setFormula('=D13-D14');
  sh.getRange(15,5).setFormula('=IF(E13="","",E13-E14)');
  // 累計（＝合計）
  sh.getRange(16,2).setFormula('=B13');
  sh.getRange(16,3).setFormula('=C13');
  sh.getRange(16,4).setFormula('=D13');
  sh.getRange(16,5).setFormula('=IF(E13="","",E13)');
  // ログ見出し
  sh.getRange(LOG_HEAD_ROW,1,1,LOG_HEADER.length).setValues([LOG_HEADER])
    .setFontWeight('bold').setBackground('#f7c6d4').setHorizontalAlignment('center');
  sh.setFrozenRows(LOG_HEAD_ROW);
  try{ sh.setColumnWidth(1,74); sh.setColumnWidth(2,210); sh.setColumnWidth(6,160); }catch(e){}
  try{ sh.getRange('A1').setNote('日付: '+date+' / '+(phaseOf_(date)||'期間外')); }catch(e){}
}

/* ===== 食事1件を追記 ===== */
function addEntry_(d){
  var sh=getOrCreateDayTab_(d.date);
  sh.appendRow([ d.time||'', d.food||'', bn_(d.kcal), bn_(d.protein), bn_(d.water), d.memo||'' ]);
  return objAssign_({ok:true}, getDay_(d.date));
}

/* ===== 朝の計測（体組成）を保存 ===== */
function saveBody_(d){
  var sh=getOrCreateDayTab_(d.date);
  var b=d.body||{};
  var prev=readBody_(findTab_(prevDate_(d.date)));
  var base=readBody_(findTab_('2026-08-28'));
  for(var i=0;i<BODY.length;i++){
    var key=BODY[i][0], r=BODY_START+i, goal=BODY[i][2];
    var cur=numOrNull_(b[key]);
    sh.getRange(r,3).setValue(cur===null?'':cur);                                  // 当日 C
    var pv=prev?prev[key]:null;
    sh.getRange(r,2).setValue(pv===null||pv===undefined?'':pv);                     // 前日 B
    sh.getRange(r,4).setValue((cur!==null&&pv!=null&&pv!=='')?round2_(cur-pv):'');   // 前日比 D
    var bv=base?base[key]:null;
    sh.getRange(r,5).setValue((cur!==null&&bv!=null&&bv!=='')?round2_(cur-bv):'');    // 当初比 E
    sh.getRange(r,7).setValue((cur!==null&&goal!=='')?round2_(goal-cur):'');          // 差分(残り) G
  }
  if(d.sleep!==undefined && d.sleep!=='' && d.sleep!==null) sh.getRange(SLEEP_CELL[0],SLEEP_CELL[1]).setValue(bn_(d.sleep));
  return objAssign_({ok:true}, getDay_(d.date));
}
function readBody_(sh){
  if(!sh) return null;
  var vals=sh.getRange(BODY_START,3,BODY.length,1).getValues();
  var o={};
  for(var i=0;i<BODY.length;i++){ var v=vals[i][0]; o[BODY[i][0]]=(v===''?null:Number(v)); }
  return o;
}

/* ===== その日の一覧＋合計＋体組成 ===== */
function getDay_(date){
  var sh=findTab_(date);
  var out={ ok:true, date:date, tab:tabName_(date), title:tabTitle_(date), rows:[], total:{kcal:0,protein:0,water:0}, count:0, body:{}, sleep:'' };
  if(!sh) return out;
  var last=sh.getLastRow();
  if(last>=LOG_START){
    var vals=sh.getRange(LOG_START,1,last-LOG_START+1,LOG_HEADER.length).getValues();
    for(var i=0;i<vals.length;i++){
      var r=vals[i];
      if(String(r.join(''))==='') continue;
      out.rows.push({ row:LOG_START+i, time:fmtCell_(r[0]), food:r[1], kcal:n_(r[2]), protein:n_(r[3]), water:n_(r[4]), memo:r[5] });
      out.total.kcal+=n_(r[2]); out.total.protein+=n_(r[3]); out.total.water+=n_(r[4]);
    }
  }
  out.total.protein=Math.round(out.total.protein*10)/10;
  out.count=out.rows.length;
  var bvals=sh.getRange(BODY_START,3,BODY.length,1).getValues();
  for(var j=0;j<BODY.length;j++){ var v=bvals[j][0]; out.body[BODY[j][0]]=(v===''?'':Number(v)); }
  var sl=sh.getRange(SLEEP_CELL[0],SLEEP_CELL[1]).getValue();
  out.sleep=(sl===''?'':Number(sl));
  return out;
}

/* ===== 1件削除 ===== */
function deleteEntry_(d){
  var sh=findTab_(d.date);
  if(sh && d.row>=LOG_START && d.row<=sh.getLastRow()) sh.deleteRow(d.row);
  return objAssign_({ok:true}, getDay_(d.date));
}

/* ===== 8/28〜9/27 全タブを（再）生成 ===== */
function setupTabs_(){
  var made=[];
  for(var i=0;i<31;i++){
    var dt=new Date(2026,7,28+i,12,0,0);
    var date=Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    var sh=getOrCreateDayTab_(date);
    buildTemplate_(sh, date);
    made.push(tabName_(date));
  }
  return { ok:true, rebuilt:made.length, tabs:made };
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
function numOrNull_(v){ if(v===undefined||v===null||v==='')return null; var n=parseFloat(v); return isNaN(n)?null:n; }
function round2_(v){ return Math.round(n_(v)*100)/100; }
function pad2_(n){ return (n<10?'0':'')+n; }
function prevDate_(date){ var p=date.split('-'); var d=new Date(Number(p[0]),Number(p[1])-1,Number(p[2])); d.setDate(d.getDate()-1); return d.getFullYear()+'-'+pad2_(d.getMonth()+1)+'-'+pad2_(d.getDate()); }
function fmtCell_(v){ if(v instanceof Date){ return Utilities.formatDate(v, TZ, 'H:mm'); } return v; }
function objAssign_(a,b){ for(var k in b){ if(b.hasOwnProperty(k)) a[k]=b[k]; } return a; }

/* ===== 動作テスト（任意）===== */
function testV3(){
  setupTabs_();
  Logger.log(addEntry_({date:'2026-08-29', time:'8:00', food:'テスト', kcal:100, protein:5, water:200, memo:''}));
  Logger.log(saveBody_({date:'2026-08-29', body:{weight:60.2,fat:31.8,bmi:22.4,water:46.7,visc:5,sub:29.5,muscle:39.7}, sleep:7}));
  Logger.log(getDay_('2026-08-29'));
}
