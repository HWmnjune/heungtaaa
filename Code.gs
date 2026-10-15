// ===== 흉터 기록 서버 (시트 v2) =====
// 게임(index.html)이 부르는 함수: logBatch · backupSave · backupLoad
// 시각 칸에는 시간만(HH:mm:ss) 적는다. 날짜는 날짜 칸에 따로 있다.
const SID='11pq1kduA7iPPPfbbACg1ZGevyQpfgexg7WcUCECxN7w';
const REG=['들판','숲','폐허','고개','바닷가','동굴','섬'];
const BASE=[['시각','tm'],['날짜','dy'],['사용자','u'],['버전','v'],['기록 ID','e']];
const RUNC=[['판 ID','r'],['몇 번째 판','rn'],['지역','rg']];
const KIND={
 open:{sh:'접속',c:[['처음 온 날','first'],['새 사용자','fresh'],['지난 접속 뒤 일수','gap'],['누적 판 수','runs'],['휴대폰','phone'],['화면 크기','scr'],['효과 줄이기','rfx'],['음악','vm'],['효과음','vs'],['이어하기 있음','save'],['저장 가능','stor']]},
 run:{sh:'판',run:1,c:[['결과','res'],['이긴 결투','fw'],['결투 시간(초)','sec'],['플레이 시간(초)','wall'],['각성','aw'],['버팀','en'],['연속 처치','c2'],['흉터','sc'],['동료','pet'],['친밀','plv'],['무기','wp'],['강화','wl'],['수련 공/막/피','tr']]},
 fight:{sh:'결투',run:1,c:[['칸','cell'],['종류','kind'],['상대','foe'],['결과','res'],['시작 체력','hp0'],['끝 체력','hp'],['최대 체력','mhp'],['상대 남은 체력(%)','ehp'],['시간(초)','sec'],['강타','hv'],['완벽 강타','pf'],['받아치기','pa'],['피하기','dg'],['막기','bl'],['맞음','ht'],['각성','aw'],['버팀','en'],['머리 흉터','gh'],['몸통 흉터','gt'],['팔 흉터','ga'],['다리 흉터','gl'],['쓰러진 부위','lp'],['동료','pet'],['무기','wp'],['강화','wl']]},
 map:{sh:'지도',run:1,c:[['칸','cell'],['구분','kind'],['장소','at'],['고른 것','pick'],['갈 수 있던 칸','opts']]},
 err:{sh:'오류',c:[['내용','msg'],['위치','at'],['화면','ph']]}};

function doGet(){
 return HtmlService.createHtmlOutputFromFile('index').setTitle('흉터')
  .addMetaTag('viewport','width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no,viewport-fit=cover')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}

function book_(){try{const a=SpreadsheetApp.getActiveSpreadsheet();if(a)return a;}catch(e){}return SpreadsheetApp.openById(SID);}
function now_(ss){const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),d=new Date();return{dy:Utilities.formatDate(d,tz,'yyyy-MM-dd'),hms:Utilities.formatDate(d,tz,'HH:mm:ss')};}
function frac_(hms){const p=String(hms).split(':').map(Number);return(p[0]*3600+p[1]*60+(p[2]||0))/86400;}
function defs_(K){return BASE.concat(K.run?RUNC:[],K.c);}
function sheet_(ss,name,heads){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
 if(sh.getLastColumn()<1){sh.getRange(1,1,1,heads.length).setValues([heads]).setFontWeight('bold');sh.setFrozenRows(1);}return sh;}
function safe_(v){if(typeof v!=='string')return v;v=v.slice(0,500);return/^[=+\-@]/.test(v)?"'"+v:v;}
function val_(o,key,T){switch(key){
 case'tm':return frac_(T.hms);
 case'dy':return o.dy||T.dy;
 case'rg':return REG[o.reg]!=null?REG[o.reg]:(o.reg==null?'':o.reg);
 case'cell':return typeof o.ly==='number'&&o.ly>=0?o.ly+1:'';
 case'tr':return o.tr?"'"+o.tr:'';}
 const v=o[key];return v==null?'':typeof v==='object'?JSON.stringify(v):v;}
function bad_(o,why){return{k:'err',e:String(o&&o.e||'x')+'.s',u:o&&o.u||'',v:o&&o.v||'',dy:o&&o.dy,msg:'서버 · '+why,at:String(o&&o.k||''),ph:''};}

// 기록 하나가 잘못돼도 나머지는 그대로 적는다
function logBatch(list){
 if(!Array.isArray(list)||!list.length)return{ok:0};
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  const ss=book_(),T=now_(ss),cache=CacheService.getScriptCache(),key=o=>'e:'+String(o.e).slice(0,80),
   ids=list.filter(o=>o&&o.e).map(key),seen=ids.length?cache.getAll(ids):{},put={},grp={},errs=[];
  list.forEach(o=>{if(!o||!o.e)return errs.push(bad_(o,'기록 ID 없음'));const id=key(o);if(seen[id]||put[id])return;
   if(!KIND[o.k])return errs.push(bad_(o,'모르는 종류 '+o.k));put[id]='1';(grp[o.k]=grp[o.k]||[]).push(o);});
  const write=(k,arr)=>{const K=KIND[k],D=defs_(K),sh=sheet_(ss,K.sh,D.map(d=>d[0])),w=sh.getLastColumn(),
    H=sh.getRange(1,1,1,w).getValues()[0].map(String),M={};D.forEach(d=>M[d[0]]=d[1]);const rows=[];
   arr.forEach(o=>{try{rows.push(H.map(h=>M[h]?safe_(val_(o,M[h],T)):''));}catch(e){if(k!=='err')errs.push(bad_(o,e.message));}});
   if(!rows.length)return;const r0=sh.getLastRow()+1;sh.getRange(r0,1,rows.length,w).setValues(rows);
   const ti=H.indexOf('시각');if(ti>=0)sh.getRange(r0,ti+1,rows.length,1).setNumberFormat('HH:mm:ss');};
  for(const k in grp){if(k==='err')continue;try{write(k,grp[k]);}catch(e){grp[k].forEach(o=>{delete put[key(o)];errs.push(bad_(o,e.message));});}}
  const ea=(grp.err||[]).concat(errs);if(ea.length)write('err',ea);
  if(Object.keys(put).length)cache.putAll(put,21600);
  return{ok:list.length};
 }finally{lock.releaseLock();}}

function find_(sh,k){const n=sh.getLastRow();if(n<2)return 0;const c=sh.getRange(2,1,n-1,1).createTextFinder(k).matchEntireCell(true).findNext();return c?c.getRow():0;}
function backupSave(u,s){u=String(u||'');s=String(s||'');if(!/^[a-z0-9]{8}$/.test(u)||!s||s.length>49000)return false;
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{const ss=book_(),T=now_(ss),sh=sheet_(ss,'백업',['사용자','맡긴 시각','버전','크기','데이터']);let v='';try{v=JSON.parse(s).v||'';}catch(e){}
  const row=[u,T.dy+' '+T.hms,v,s.length,s],r=find_(sh,u);if(r)sh.getRange(r,1,1,5).setValues([row]);else sh.appendRow(row);return true;
 }finally{lock.releaseLock();}}
function backupLoad(code){code=String(code||'').toLowerCase().replace(/[^a-z0-9]/g,'');if(!/^[a-z0-9]{8}$/.test(code))return'';const ss=book_();
 for(const [n,k] of [['백업',code],['백업 (이전)','c:'+code]]){const sh=ss.getSheetByName(n);if(!sh)continue;const r=find_(sh,k);if(r)return String(sh.getRange(r,5).getValue()||'');}
 return'';}

// 한 번만 실행: 이미 적힌 시각 칸에서 날짜를 지우고 시간만 남긴다
function fixTimes(){const ss=book_(),tz=ss.getSpreadsheetTimeZone();
 ['접속','판','결투','지도','오류'].forEach(n=>{const sh=ss.getSheetByName(n);if(!sh||sh.getLastRow()<2)return;
  const H=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),ti=H.indexOf('시각');if(ti<0)return;
  const g=sh.getRange(2,ti+1,sh.getLastRow()-1,1),v=g.getValues().map(r=>{const x=r[0];
   if(x instanceof Date)return[frac_(Utilities.formatDate(x,tz,'HH:mm:ss'))];const m=String(x).match(/(\d{1,2}):(\d{2}):(\d{2})/);return[m?frac_(m[0]):x];});
  g.setValues(v).setNumberFormat('HH:mm:ss');});}
