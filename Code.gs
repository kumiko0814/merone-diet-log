/*************************************************************
 *  Daily Diet Log v4  →  各日タブ＝元テンプレートを完全再現
 *  ------------------------------------------------------------
 *  貼り方（更新も同じ）
 *   1. スプレッドシート → 拡張機能 → Apps Script
 *   2. 全部貼り付けて保存
 *   3. デプロイ → デプロイを管理 → 鉛筆 → バージョン「新バージョン」→ デプロイ
 *
 *  タブの構成（元の減量期Day1と同じ）:
 *   1行目 : A=タイトル / B=日付 / H=氏名(ラベル) I=氏名(値)
 *   3-5行 : H=排便有無/便秘薬摂取/体調 (ラベル) I=値
 *   6行目 : 緑ヘッダー 時刻/食事/カロリー/プロテイン/水/睡眠時間/入浴/有酸素運動/マッサージ
 *   7-36行: 記録エリア(30行・1件=1行で埋める)
 *   37行 : 食事回数
 *   38-41: 合計/理想/差分/累計
 *   43-50: 体組成(前日/当日/前日比/当初比/目標/差分)  オレンジ見出し
 *   51行 : Day0との比較 注記
 *************************************************************/

var TOKEN = 'merone-diet-2026';
var TZ = 'Asia/Tokyo';
var BUG_SHEET = '不具合ログ';

// レイアウト行
var HDR_ROW=6, LOG_TOP=7, LOG_BOT=36;
var CNT_ROW=37, SUM_ROW=38, IDEAL_ROW=39, DIFF_ROW=40, CUM_ROW=41;
var BODY_HDR=43, BODY_TOP=44, NOTE_ROW=51;

var LOG_HDR = ['時刻','食事','カロリー(kcal)','プロテイン(g)','水(ml)','睡眠時間','入浴(m)','有酸素運動(m)','マッサージ(m)'];
// 体組成 [label, 目標]   行44〜50 / 当日=C列
var BODY = [
  ['体重(kg)','weight',47],
  ['体脂肪(%)','fat',22],
  ['BMI','bmi',23],
  ['水分率(%)','water',57],
  ['内臓脂肪','visc',7],
  ['皮下脂肪','sub',19],
  ['骨格筋(%)','muscle','']
];

/* ===== entry points ===== */
function doGet(e){
  try{
    var p=(e&&e.parameter)||{};
    if(p.date) return json(getDay_(p.date));
    return json({ok:true, name:'merone-diet4', title:ss_().getName()});
  }catch(err){ return json({ok:false, error:String(err)}); }
}
function doPost(e){
  try{
    var d=JSON.parse(e.postData.contents);
    if(d.token!==TOKEN) return json({ok:false, error:'あいことば(token)が違います'});
    switch(d.action){
      case 'addEntry':    return json(addEntry_(d));
      case 'saveDaily':   return json(saveDaily_(d));
      case 'getDay':      return json(getDay_(d.date));
      case 'deleteEntry': return json(deleteEntry_(d));
      case 'setupTabs':   return json(setupTabs_());
      case 'bug':         logBug_(d); return json({ok:true});
      default:            return json({ok:false, error:'unknown action: '+d.action});
    }
  }catch(err){ return json({ok:false, error:String(err)}); }
}

/* ===== 期・タブ名・日付 ===== */
function phaseOf_(date){
  if(date>='2026-08-28' && date<='2026-08-30') return '分析期';
  if(date>='2026-08-31' && date<='2026-09-10') return '減量期';
  if(date>='2026-09-11' && date<='2026-09-27') return '分析期';
  return '';
}
function tabName_(date){ var p=String(date).split('-'); var nm=Number(p[1])+'/'+Number(p[2]); var ph=phaseOf_(date); return ph?nm+' '+ph:nm; }
function tabTitle_(date){
  var p=String(date).split('-'); var d=new Date(Number(p[0]),Number(p[1])-1,Number(p[2]));
  var dow=['日','月','火','水','木','金','土'][d.getDay()]; var ph=phaseOf_(date);
  return Number(p[1])+'月'+Number(p[2])+'日（'+dow+'）'+(ph?'　'+ph:'');
}
function dotDate_(date){ var p=String(date).split('-'); return Number(p[0])+'.'+Number(p[1])+'.'+Number(p[2]); }
function findTab_(date){ var ss=ss_(),n=tabName_(date); return ss.getSheetByName(n)||ss.getSheetByName(n.replace(/\//g,'-')); }
function getOrCreateDayTab_(date){
  var sh=findTab_(date); if(sh) return sh;
  var ss=ss_(), name=tabName_(date), c;
  try{ c=ss.insertSheet(name); }catch(e){ name=name.replace(/\//g,'-'); c=ss.getSheetByName(name)||ss.insertSheet(name); }
  buildTemplate_(c, date); return c;
}
function colL_(c){ return String.fromCharCode(64+c); }   // 1->A

/* ===== テンプレート（元の完全再現） ===== */
function buildTemplate_(sh, date){
  sh.clear();
  // 1行目
  sh.getRange(1,1).setValue(tabTitle_(date)).setFontWeight('bold');
  sh.getRange(1,2).setValue(dotDate_(date));
  sh.getRange(1,8).setValue('氏名').setFontWeight('bold').setHorizontalAlignment('right');
  // 3-5 ラベル
  sh.getRange(3,8).setValue('排便有無').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(4,8).setValue('便秘薬摂取').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(5,8).setValue('体調').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(1,9,5,1).setBorder(true,true,true,true,false,false);   // 氏名〜体調の値枠
  // 緑ヘッダー
  sh.getRange(HDR_ROW,1,1,9).setValues([LOG_HDR]).setFontWeight('bold').setBackground('#8cff6b').setHorizontalAlignment('center');
  // 記録エリア枠
  sh.getRange(HDR_ROW,1,LOG_BOT-HDR_ROW+1,9).setBorder(true,true,true,true,true,true);
  // 食事回数
  sh.getRange(CNT_ROW,1).setValue('→食事回数：').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(CNT_ROW,2).setFormula('=COUNTA(B'+LOG_TOP+':B'+LOG_BOT+')');
  // 合計/理想/差分/累計 ラベル
  sh.getRange(SUM_ROW,2).setValue('合計').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(IDEAL_ROW,2).setValue('理想').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(DIFF_ROW,2).setValue('差分').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(CUM_ROW,2).setValue('累計').setFontWeight('bold').setHorizontalAlignment('right');
  // 合計: C/D/E はログ自動集計、F〜I(睡眠/入浴/有酸素/マッサージ)は入力値
  sh.getRange(SUM_ROW,3).setFormula('=SUM(C'+LOG_TOP+':C'+LOG_BOT+')');
  sh.getRange(SUM_ROW,4).setFormula('=SUM(D'+LOG_TOP+':D'+LOG_BOT+')');
  sh.getRange(SUM_ROW,5).setFormula('=SUM(E'+LOG_TOP+':E'+LOG_BOT+')');
  // 理想 C〜I
  sh.getRange(IDEAL_ROW,3,1,7).setValues([[470,80,3000,7,60,60,60]]);
  // 差分 C〜I = 合計-理想
  for(var c=3;c<=9;c++) sh.getRange(DIFF_ROW,c).setFormula('=IF('+colL_(c)+SUM_ROW+'="","",'+colL_(c)+SUM_ROW+'-'+colL_(c)+IDEAL_ROW+')');
  // 累計 = 合計
  for(var c2=3;c2<=9;c2++) sh.getRange(CUM_ROW,c2).setFormula('='+colL_(c2)+SUM_ROW);
  sh.getRange(SUM_ROW,2,4,8).setBorder(true,true,true,true,true,true);
  // 体組成ヘッダー
  sh.getRange(BODY_HDR,2,1,6).setValues([['前日','当日','前日比','当初比*','目標','差分(残り)']]).setFontWeight('bold').setHorizontalAlignment('center');
  for(var i=0;i<BODY.length;i++){
    var r=BODY_TOP+i;
    sh.getRange(r,1).setValue(BODY[i][0]).setFontWeight('bold').setBackground('#f6a04b').setHorizontalAlignment('center');
    if(BODY[i][2]!=='') sh.getRange(r,6).setValue(BODY[i][2]);   // 目標=F
  }
  sh.getRange(BODY_HDR,1,BODY.length+1,7).setBorder(true,true,true,true,true,true);
  sh.getRange(NOTE_ROW,5).setValue('*Day 0との比較');
  // 体裁
  sh.setFrozenRows(HDR_ROW);
  try{
    sh.setColumnWidth(1,90); sh.setColumnWidth(2,170);
    for(var w=3;w<=9;w++) sh.setColumnWidth(w,108);
  }catch(e){}
  try{ sh.getRange('A1').setNote('日付: '+date+' / '+(phaseOf_(date)||'期間外')); }catch(e){}
}

/* ===== 食事1件を記録エリアに追記 ===== */
function addEntry_(d){
  var sh=getOrCreateDayTab_(d.date);
  var r=firstEmptyLog_(sh);
  sh.getRange(r,1,1,5).setValues([[ d.time||'', d.food||'', bn_(d.kcal), bn_(d.protein), bn_(d.water) ]]);
  return objAssign_({ok:true}, getDay_(d.date));
}
function firstEmptyLog_(sh){
  var vals=sh.getRange(LOG_TOP,1,LOG_BOT-LOG_TOP+1,2).getValues();
  for(var i=0;i<vals.length;i++){ if(String(vals[i][0])==='' && String(vals[i][1])===''){ return LOG_TOP+i; } }
  return LOG_BOT;   // 満杯時は最終行に（通常30件で足りる）
}

/* ===== 1日の からだ・ケア を保存 ===== */
function saveDaily_(d){
  var sh=getOrCreateDayTab_(d.date);
  // 上部ラベル値
  if(d.name!==undefined)   sh.getRange(1,9).setValue(d.name||'');
  if(d.haiben!==undefined) sh.getRange(3,9).setValue(d.haiben||'');
  if(d.benpi!==undefined)  sh.getRange(4,9).setValue(d.benpi||'');
  if(d.taicho!==undefined) sh.getRange(5,9).setValue(d.taicho||'');
  // 合計行の 睡眠/入浴/有酸素/マッサージ（F〜I=値）
  if(d.sleep!==undefined)   sh.getRange(SUM_ROW,6).setValue(bn_(d.sleep));
  if(d.bath!==undefined)    sh.getRange(SUM_ROW,7).setValue(bn_(d.bath));
  if(d.aerobic!==undefined) sh.getRange(SUM_ROW,8).setValue(bn_(d.aerobic));
  if(d.massage!==undefined) sh.getRange(SUM_ROW,9).setValue(bn_(d.massage));
  if(d.sleep!==undefined && d.sleep!=='' && d.sleep!==null) sh.getRange(CNT_ROW,6).setValue(bn_(d.sleep)+'時間');
  // 体組成
  var b=d.body||{};
  var prev=readBody_(findTab_(prevDate_(d.date)));
  var base=readBody_(findTab_('2026-08-28'));
  for(var i=0;i<BODY.length;i++){
    var key=BODY[i][1], goal=BODY[i][2], r=BODY_TOP+i;
    var cur=numOrNull_(b[key]);
    sh.getRange(r,3).setValue(cur===null?'':cur);                                    // 当日 C
    var pv=prev?prev[key]:null;
    sh.getRange(r,2).setValue(pv===null||pv===undefined?'':pv);                       // 前日 B
    sh.getRange(r,4).setValue((cur!==null&&pv!=null&&pv!=='')?round2_(cur-pv):'');     // 前日比 D
    var bv=base?base[key]:null;
    sh.getRange(r,5).setValue((cur!==null&&bv!=null&&bv!=='')?round2_(cur-bv):'');      // 当初比 E
    sh.getRange(r,7).setValue((cur!==null&&goal!=='')?round2_(goal-cur):'');            // 差分 G
  }
  return objAssign_({ok:true}, getDay_(d.date));
}
function readBody_(sh){
  if(!sh) return null;
  var vals=sh.getRange(BODY_TOP,3,BODY.length,1).getValues(); var o={};
  for(var i=0;i<BODY.length;i++){ var v=vals[i][0]; o[BODY[i][1]]=(v===''?null:Number(v)); }
  return o;
}

/* ===== その日の一覧＋合計＋からだ ===== */
function getDay_(date){
  var sh=findTab_(date);
  var out={ ok:true, date:date, tab:tabName_(date), title:tabTitle_(date), rows:[], total:{kcal:0,protein:0,water:0}, count:0,
            body:{}, name:'', haiben:'', benpi:'', taicho:'', sleep:'', bath:'', aerobic:'', massage:'' };
  if(!sh) return out;
  var vals=sh.getRange(LOG_TOP,1,LOG_BOT-LOG_TOP+1,9).getValues();
  for(var i=0;i<vals.length;i++){
    var r=vals[i];
    if(String(r[0])==='' && String(r[1])==='') continue;
    out.rows.push({ row:LOG_TOP+i, time:fmtCell_(r[0]), food:r[1], kcal:n_(r[2]), protein:n_(r[3]), water:n_(r[4]) });
    out.total.kcal+=n_(r[2]); out.total.protein+=n_(r[3]); out.total.water+=n_(r[4]);
  }
  out.total.protein=Math.round(out.total.protein*10)/10; out.count=out.rows.length;
  var bvals=sh.getRange(BODY_TOP,3,BODY.length,1).getValues();
  for(var j=0;j<BODY.length;j++){ var v=bvals[j][0]; out.body[BODY[j][1]]=(v===''?'':Number(v)); }
  out.name  = sh.getRange(1,9).getValue();
  out.haiben= sh.getRange(3,9).getValue();
  out.benpi = sh.getRange(4,9).getValue();
  out.taicho= sh.getRange(5,9).getValue();
  out.sleep   = cellNum_(sh.getRange(SUM_ROW,6).getValue());
  out.bath    = cellNum_(sh.getRange(SUM_ROW,7).getValue());
  out.aerobic = cellNum_(sh.getRange(SUM_ROW,8).getValue());
  out.massage = cellNum_(sh.getRange(SUM_ROW,9).getValue());
  return out;
}

/* ===== 1件削除（行を空にする＝レイアウト維持） ===== */
function deleteEntry_(d){
  var sh=findTab_(d.date);
  if(sh && d.row>=LOG_TOP && d.row<=LOG_BOT) sh.getRange(d.row,1,1,9).clearContent();
  return objAssign_({ok:true}, getDay_(d.date));
}

/* ===== 8/28〜9/27 全タブを（再）生成 ===== */
function setupTabs_(){
  var made=[];
  for(var i=0;i<31;i++){
    var dt=new Date(2026,7,28+i,12,0,0);
    var date=Utilities.formatDate(dt,TZ,'yyyy-MM-dd');
    var sh=getOrCreateDayTab_(date);
    buildTemplate_(sh,date);
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
function cellNum_(v){ return (v===''||v===null||v===undefined)?'':(isNaN(Number(v))?v:Number(v)); }
function numOrNull_(v){ if(v===undefined||v===null||v==='')return null; var x=parseFloat(v); return isNaN(x)?null:x; }
function round2_(v){ return Math.round(n_(v)*100)/100; }
function pad2_(n){ return (n<10?'0':'')+n; }
function prevDate_(date){ var p=date.split('-'); var d=new Date(Number(p[0]),Number(p[1])-1,Number(p[2])); d.setDate(d.getDate()-1); return d.getFullYear()+'-'+pad2_(d.getMonth()+1)+'-'+pad2_(d.getDate()); }
function fmtCell_(v){ if(v instanceof Date){ return Utilities.formatDate(v,TZ,'H:mm'); } return v; }
function objAssign_(a,b){ for(var k in b){ if(b.hasOwnProperty(k)) a[k]=b[k]; } return a; }

/* ===== 動作テスト（任意）===== */
function testV4(){
  setupTabs_();
  Logger.log(addEntry_({date:'2026-08-29', time:'8:00', food:'テスト', kcal:100, protein:5, water:200}));
  Logger.log(saveDaily_({date:'2026-08-29', name:'森川', haiben:'あり', benpi:'なし', taicho:'良好',
    sleep:7, bath:60, aerobic:30, massage:20,
    body:{weight:60.2,fat:31.8,bmi:22.4,water:46.7,visc:5,sub:29.5,muscle:39.7}}));
  Logger.log(getDay_('2026-08-29'));
}
