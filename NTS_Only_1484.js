/** 국세청 휴폐업 전용 1,484건 독립 조회. DART/G2B/점수화/입력삭제 실행 안 함. */
const NTS_ONLY={source:'대량분석_신규입력',prefix:'휴폐업조회_',fn:'ntsOnlyTick',size:100};
function ntsOnlyStart(){
 const ss=SpreadsheetApp.openById(NIGHT.ssid),src=ss.getSheetByName(NTS_ONLY.source);
 if(!src||src.getLastRow()<2)throw Error('입력 자료 없음');
 if(!nightNtsKey_())throw Error('NTS_API_KEY 미설정');
 const lock=LockService.getScriptLock();if(!lock.tryLock(15000))throw Error('다른 작업 실행 중');
 try{
  const p=PropertiesService.getScriptProperties();
  if(p.getProperty('NTS_ONLY_ACTIVE')==='1')return '휴폐업 조회 이미 진행 중';
  const input=src.getRange(2,1,src.getLastRow()-1,7).getDisplayValues().filter(r=>r[1]||r[3]);
  if(!input.length)throw Error('입력 기업 없음');
  const run=Utilities.formatDate(new Date(),'Asia/Seoul','yyyyMMdd-HHmmss'),title=NTS_ONLY.prefix+run;
  const headers=['번호','기업명','사업자등록번호','조회상태','사업자상태','과세유형','폐업일자','조회시각','비고'];
  const rows=input.map((r,i)=>[i+1,r[1],String(r[3]).replace(/\D/g,''),'대기','','','','','']);
  const sh=ss.insertSheet(title);sh.getRange(1,1,1,headers.length).setValues([headers]);
  for(let i=0;i<rows.length;i+=300)sh.getRange(i+2,1,Math.min(300,rows.length-i),headers.length).setValues(rows.slice(i,i+300));
  sh.setFrozenRows(1);sh.getRange(1,1,1,9).setBackground('#173451').setFontColor('#ffffff').setFontWeight('bold');
  sh.setColumnWidth(2,210);sh.setColumnWidth(5,180);
  p.setProperties({NTS_ONLY_ACTIVE:'1',NTS_ONLY_SHEET:title,NTS_ONLY_COUNT:String(rows.length)});
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NTS_ONLY.fn).forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(NTS_ONLY.fn).timeBased().everyMinutes(5).create();
  return {sheet:title,total:rows.length,scheduled:true};
 }finally{lock.releaseLock();}
}
function ntsOnlyTick(){
 const p=PropertiesService.getScriptProperties();if(p.getProperty('NTS_ONLY_ACTIVE')!=='1')return '비활성';
 const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return '다른 작업 중';
 try{
  const ss=SpreadsheetApp.openById(NIGHT.ssid),sh=ss.getSheetByName(p.getProperty('NTS_ONLY_SHEET'));
  if(!sh)throw Error('조회 결과 시트 없음');
  const data=sh.getRange(2,1,sh.getLastRow()-1,9).getDisplayValues(),todo=[];
  for(let i=0;i<data.length&&todo.length<NTS_ONLY.size;i++)if(data[i][3]==='대기')todo.push({row:i+2,biz:String(data[i][2]).replace(/\D/g,'')});
  if(!todo.length){ntsOnlyStop();return '전수 조회 처리 종료';}
  const invalid=todo.filter(x=>!/^\d{10}$/.test(x.biz));
  for(const x of invalid)sh.getRange(x.row,4,1,6).setValues([['확인필요','','','',new Date(),'사업자번호 형식 오류']]);
  const valid=todo.filter(x=>/^\d{10}$/.test(x.biz));
  if(!valid.length)return '형식 오류 '+invalid.length+'건';
  const url='https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey='+encodeURIComponent(nightNtsKey_());
  let result;
  try{
   const res=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify({b_no:valid.map(x=>x.biz)}),muteHttpExceptions:true});
   if(res.getResponseCode()!==200)throw Error('국세청 HTTP '+res.getResponseCode());
   result=JSON.parse(res.getContentText());if(!Array.isArray(result.data))throw Error('응답 형식 오류');
  }catch(e){
   // 조회 실패 시 입력 보존. 다음 트리거에서 재시도; 인증키·URL은 기록하지 않음.
   console.log('휴폐업 API 오류(인증정보 비표시): '+String(e.message||'').replace(/serviceKey=[^&\s]+/g,'serviceKey=[REDACTED]').slice(0,120));
   return 'API 오류, 다음 회차 재시도';
  }
  const m=new Map(result.data.map(x=>[String(x.b_no||'').replace(/\D/g,''),x]));
  for(const x of valid){
   const v=m.get(x.biz),code=String(v?.b_stt_cd||'');
   const state=code==='01'?'계속사업':code==='02'?'휴업':code==='03'?'폐업':'확인필요';
   sh.getRange(x.row,4,1,6).setValues([[v?'조회완료':'확인필요',v?state+' / '+String(v.b_stt||''):'응답 누락',String(v?.tax_type||''),String(v?.end_dt||''),new Date(),v?'국세청 상태조회':'해당 사업자번호 응답 없음']]);
  }
  return {processed:valid.length,remaining:data.filter(r=>r[3]==='대기').length-todo.length,sheet:sh.getName()};
 }finally{lock.releaseLock();}
}
function ntsOnlyStop(){
 const p=PropertiesService.getScriptProperties();p.setProperty('NTS_ONLY_ACTIVE','0');
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NTS_ONLY.fn).forEach(t=>ScriptApp.deleteTrigger(t));
 return '휴폐업 전용 트리거 중지(결과 보존)';
}
