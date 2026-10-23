/**
 * 당일 기업 1,000건과 나라장터 계약정보 대조.
 * 계약 체결 사실은 미수채권 존재·연체의 증거가 아닙니다.
 * 첫 실행: dailyG2BStart('20261022','20261022') (검색할 계약체결일 범위 지정)
 * 진행: dailyG2BTick (10분 자동), 상태: dailyG2BStatus()
 * 일일_분석이력과 일일_배포대기에는 일치 건만 별도 검증 표시.
 */
const DAILY_G2B={sheet:'일일_나라장터매칭',state:'DAILY_G2B_STATE',fn:'dailyG2BTick',batch:6,
 ops:[['물품','getCntrctInfoListThng'],['공사','getCntrctInfoListCnstwk'],['용역','getCntrctInfoListServc'],['외자','getCntrctInfoListFrgcpt']]};
function dailyG2BStatus(){const s=PropertiesService.getScriptProperties().getProperty(DAILY_G2B.state);return s?JSON.parse(s):'미시작';}
function dailyG2BStart(from,to){
 // Apps Script 함수 메뉴에서 인수 없이 실행하면 전날(한국시간) 계약정보를 조회합니다.
 // 날짜를 명시한 호출은 그대로 지원합니다.
 if(from===undefined&&to===undefined){
  const yesterday=new Date(Date.now()-86400000);
  const day=Utilities.formatDate(yesterday,'Asia/Seoul','yyyyMMdd');
  from=day;to=day;
 }
 const p=PropertiesService.getScriptProperties(),run=p.getProperty('DAILY_RUN');
 if(!run)throw Error('DAILY_RUN 없음');
 if(!p.getProperty('G2B_API_KEY'))throw Error('G2B_API_KEY 없음: Apps Script 스크립트 속성에 설정 필요');
 if(!/^\d{8}$/.test(String(from))||!/^\d{8}$/.test(String(to))||from>to)throw Error('검색기간은 YYYYMMDD, 시작일 <= 종료일');
 const old=dailyG2BStatus();
 if(old!=='미시작'&&old.status==='running')throw Error('기존 나라장터 작업 진행 중');
 const ss=SpreadsheetApp.openById(NIGHT.ssid),q=ss.getSheetByName(NIGHT.queue);
 const targets=q.getRange(2,1,q.getLastRow()-1,16).getDisplayValues().filter(r=>r[0]===run);
 if(!targets.length)throw Error('당일 실행 대기열 없음');
 const sheet=ss.getSheetByName(DAILY_G2B.sheet)||ss.insertSheet(DAILY_G2B.sheet);
 if(sheet.getLastRow()===0)sheet.appendRow(['실행ID','조회기간','분야','사업자등록번호','기업명','계약명','계약번호','계약금액','계약일','검증수준']);
 const state={run,from,to,category:0,page:1,processed:0,matches:0,status:'running',updated:new Date().toISOString()};
 p.setProperty(DAILY_G2B.state,JSON.stringify(state));
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===DAILY_G2B.fn).forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger(DAILY_G2B.fn).timeBased().everyMinutes(10).create();
 return '나라장터 대조 예약: '+run+' / '+from+'~'+to+' / 대상 '+targets.length+'건';
}
function dailyG2BTick(){
 const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return '다른 작업 중';
 try{
  const p=PropertiesService.getScriptProperties(),state=dailyG2BStatus();
  if(state==='미시작'||state.status!=='running')return '대기';
  const ss=SpreadsheetApp.openById(NIGHT.ssid),q=ss.getSheetByName(NIGHT.queue);
  const rows=q.getRange(2,1,q.getLastRow()-1,16).getDisplayValues().filter(r=>r[0]===state.run);
  const targets=new Map(rows.map(r=>[String(r[3]).replace(/\D/g,''),r[2]]));
  const key=p.getProperty('G2B_API_KEY'),sheet=ss.getSheetByName(DAILY_G2B.sheet);
  for(let i=0;i<DAILY_G2B.batch&&state.category<DAILY_G2B.ops.length;i++){
   const [kind,op]=DAILY_G2B.ops[state.category];
   const qs={ServiceKey:key,numOfRows:100,pageNo:state.page,type:'json',inqryDiv:1,inqryBgnDt:state.from+'0000',inqryEndDt:state.to+'2359'};
   const url='https://apis.data.go.kr/1230000/ao/CntrctInfoService/'+op+'?'+Object.keys(qs).map(k=>k+'='+encodeURIComponent(qs[k])).join('&');
   const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
   let data;try{data=JSON.parse(res.getContentText());}catch(e){throw Error(kind+' '+state.page+'페이지 응답 JSON 오류');}
   const response=data.response||{},body=response.body||{},code=String((response.header||{}).resultCode);
   if(res.getResponseCode()!==200||!['0','00'].includes(code))throw Error(kind+' '+state.page+'페이지 HTTP '+res.getResponseCode()+' API '+code);
   let items=Array.isArray(body.items)?body.items:(body.items&&body.items.item)||[];
   if(!Array.isArray(items))items=items?[items]:[];
   const found=[];
   items.forEach(item=>{
    const nums=[...new Set(String(item.corpList||'').match(/\d{10}/g)||[])];
    nums.forEach(no=>{if(targets.has(no))found.push([state.run,state.from+'~'+state.to,kind,no,targets.get(no),String(item.cntrctNm||''),String(item.untyCntrctNo||''),String(item.totCntrctAmt||''),String(item.cntrctCnclsDate||''),'계약업체 번호 일치 / 미수채권 미확인']);});
   });
   if(found.length){if(sheet.getMaxRows()<sheet.getLastRow()+found.length)sheet.insertRowsAfter(sheet.getMaxRows(),sheet.getLastRow()+found.length-sheet.getMaxRows());sheet.getRange(sheet.getLastRow()+1,1,found.length,10).setValues(found);}
   state.processed+=items.length;state.matches+=found.length;
   if(!items.length||state.page*100>=Number(body.totalCount||0)){state.category++;state.page=1;}else state.page++;
   state.updated=new Date().toISOString();p.setProperty(DAILY_G2B.state,JSON.stringify(state));
  }
  if(state.category>=DAILY_G2B.ops.length){
   state.status='complete';p.setProperty(DAILY_G2B.state,JSON.stringify(state));
   ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===DAILY_G2B.fn).forEach(t=>ScriptApp.deleteTrigger(t));
   dailyG2BApply();
  }
  return state;
 }catch(e){
  const p=PropertiesService.getScriptProperties(),state=dailyG2BStatus();
  if(state!=='미시작'){state.status='paused';state.error=String(e.message||e).replace(/https?:\/\/[^\s]+/g,'[API URL REDACTED]');p.setProperty(DAILY_G2B.state,JSON.stringify(state));}
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===DAILY_G2B.fn).forEach(t=>ScriptApp.deleteTrigger(t));throw e;
 }finally{lock.releaseLock();}
}
function dailyG2BApply(){
 const state=dailyG2BStatus();if(state==='미시작'||state.status!=='complete')throw Error('나라장터 검색이 아직 완료되지 않았습니다');
 const ss=SpreadsheetApp.openById(NIGHT.ssid),s=ss.getSheetByName(DAILY_G2B.sheet);
 const rows=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,10).getDisplayValues().filter(r=>r[0]===state.run):[];
 const count=new Map();rows.forEach(r=>count.set(r[3],(count.get(r[3])||0)+1));
 const names=['일일_배포대기','일일_분석이력'];let updated=0;
 names.forEach(name=>{
  const sh=ss.getSheetByName(name);if(!sh||sh.getLastRow()<2)return;
  const data=sh.getRange(2,1,sh.getLastRow()-1,12).getDisplayValues();
  const changes=[];
  data.forEach((r,i)=>{if(r[0]!==state.run)return;const hits=count.get(String(r[4]).replace(/\D/g,''))||0;changes.push({row:i+2,text:hits?'계약번호 일치 '+hits+'건 (미수채권 미확인)':'해당 조회기간 계약 일치 없음 (미수채권 미확인)'});});
  changes.forEach(c=>sh.getRange(c.row,9).setValue(c.text));updated+=changes.length;
 });
 dailyLog_('나라장터','실행 '+state.run+' / '+state.from+'~'+state.to+' / 계약 '+state.processed+'건 검색 / 일치 '+state.matches+'건 / 결과행 '+updated+'개 반영');
 return {run:state.run,matchedContracts:state.matches,rowsUpdated:updated,coverage:state.from+'~'+state.to};
}
function dailyG2BResume(){
 const p=PropertiesService.getScriptProperties(),s=dailyG2BStatus();if(s==='미시작')throw Error('시작된 작업 없음');
 if(s.status==='complete'||s.status==='running')return s.status;
 s.status='running';delete s.error;p.setProperty(DAILY_G2B.state,JSON.stringify(s));
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===DAILY_G2B.fn).forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger(DAILY_G2B.fn).timeBased().everyMinutes(10).create();return '재개 예약';
}