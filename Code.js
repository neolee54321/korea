/**
 * 고려신용정보 | 독립형 무료 야간분석 v1.0 (2026-10-21)
 * Google Apps Script 신규 '독립형 프로젝트'에만 설치하세요.
 * 기존 TOP10 스크립트, 트리거, 시트는 수정하지 않습니다.
 * 무료 할당량 내 분할 실행. 실제 미수채권 존재/외부위탁 가능성은 추정하지 않습니다.
 * DART OpenAPI 사용: https://opendart.fss.or.kr/
 */
const NIGHT = Object.freeze({
  ssid:'1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU',
  input:'대량분석_신규입력', queue:'야간분석_작업대기열',
  control:'야간분석_실행관리', errors:'야간분석_오류기록',
  result:'대량분석_분석결과', agent:'대량분석_관리사배분',
  tickFn:'nightTick', tickMinutes:5, maxPerTick:15,
  maxRuntimeMs:220000, maxRetries:3, reportYear:2025
});
const H = {
 queue:['실행ID','원본행','기업명','사업자등록번호','대표전화','주소','업종','처리단계','처리상태','시도횟수','다음시도시각','마지막오류','DART고유번호','완료시각','결과행','원본출처'],
 control:['실행ID','시작요청시각','실제시작시각','최종갱신','전체건수','완료','부분완료','보류','실패','진행상태','다음 실행예정','비고'],
 result:['기업명','사업자등록번호','대표전화','주소','DART 고유번호','사업자·법인 검증','재무구분(연결/별도)','2024·2025 매출·채권 변화','연체·손실충당금 근거','검증 상태','재무 검토 의견','오류·보류 사유','분석일·출처'],
 agent:['실행ID','관리사','기업명','사업자등록번호','대표전화','주소','업종','재무징후','검증수준','검증근거','영업검토 우선순위','배분상태','비고']
};
function nightBook_(){return SpreadsheetApp.openById(NIGHT.ssid);}
function nightSheet_(name, headers){const ss=nightBook_();let s=ss.getSheetByName(name);if(!s) s=ss.insertSheet(name);if(headers){if(s.getMaxColumns()<headers.length)s.insertColumnsAfter(s.getMaxColumns(),headers.length-s.getMaxColumns());s.getRange(1,1,1,headers.length).setValues([headers]);s.setFrozenRows(1);}return s;}
function nightLock_(fn){const lock=LockService.getScriptLock();if(!lock.tryLock(5000))throw Error('다른 야간분석 작업이 실행 중입니다.');try{return fn();}finally{lock.releaseLock();}}
function nightNow_(){return new Date();}
function nightDate_(d){return Utilities.formatDate(d,'Asia/Seoul','yyyy-MM-dd HH:mm:ss');}
function nightKey_(){return PropertiesService.getScriptProperties().getProperty('DART_API_KEY');}
function nightValidate_(){if(!nightKey_())throw Error('이 독립형 프로젝트의 스크립트 속성 DART_API_KEY를 먼저 등록해야 합니다. 기존 프로젝트의 키는 자동 공유되지 않습니다.');}
/** 최초 한 번: 기존 시트 구조 확인 및 실행 메뉴 생성. 기존 시트의 데이터 행은 삭제하지 않습니다. */
function nightSetup(){nightLock_(()=>{nightSheet_(NIGHT.queue,H.queue);nightSheet_(NIGHT.control,H.control);nightSheet_(NIGHT.result,H.result);nightSheet_(NIGHT.agent,H.agent);nightSheet_(NIGHT.errors,['실행ID','기업명','사업자등록번호','단계','오류유형','오류내용','발생시각','재시도결과','최종처리']);if(!nightBook_().getSheetByName(NIGHT.input))throw Error('입력 시트 없음');nightValidate_();});return '설치 준비 확인 완료';}
/** 사용자가 실행 버튼처럼 한 번 실행. 이 시점에 입력된 모든 행을 새로운 실행으로 등록합니다. 중복 필터 없음. */
/** 대량분석 전용 입출력 시트 생성. 기존 20건 시험 시트는 변경하지 않습니다. */
function nightBulkSetup(){
 return nightLock_(()=>{
  const ss=nightBook_();
  const name=NIGHT.input;
  const expected=['기업유형','기업명','전화번호','사업자번호','대표자명','기업주소','산업명'];
  let sheet=ss.getSheetByName(name);
  if(!sheet){sheet=ss.insertSheet(name);sheet.getRange(1,1,1,7).setValues([expected]);}
  else {
   const header=sheet.getRange(1,1,1,7).getDisplayValues()[0];
   if(expected.some((h,i)=>header[i]!==h))throw Error('기존 대량 입력 시트의 A:G 헤더가 다릅니다. 자동 덮어쓰기를 중단합니다.');
  }
  sheet.setFrozenRows(1);
  sheet.getRange('D:D').setNumberFormat('@');
  nightSheet_(NIGHT.result,H.result);
  nightSheet_(NIGHT.agent,H.agent);
  nightSheet_(NIGHT.queue,H.queue);
  nightSheet_(NIGHT.control,H.control);
  return '대량분석 입력·결과 시트 준비 완료 / 기존 20건 자료 유지';
 });
}
function nightStart(){return nightLock_(()=>{nightValidate_();const props=PropertiesService.getScriptProperties();if(props.getProperty('NIGHT_ACTIVE')==='1')throw Error('이미 야간분석이 진행 중입니다.');const input=nightBook_().getSheetByName(NIGHT.input);if(!input)throw Error('대량분석_신규입력 시트가 없습니다. nightBulkSetup을 먼저 실행하세요.');const expected=['기업유형','기업명','전화번호','사업자번호','대표자명','기업주소','산업명'];const header=input.getRange(1,1,1,7).getDisplayValues()[0];if(expected.some((h,i)=>header[i]!==h))throw Error('비즈라인 입력 헤더 A:G 순서가 맞지 않습니다.');const values=input.getDataRange().getDisplayValues().slice(1);const tasks=[];const id=Utilities.formatDate(new Date(),'Asia/Seoul','yyyyMMdd-HHmmss');values.forEach((v,i)=>{if(!String(v[1]||'').trim())return;tasks.push([id,i+2,v[1],v[3],v[2],v[5],v[6],'대기','대기',0,'','','','','','비즈라인']);});if(!tasks.length)throw Error('입력 기업이 없습니다.');const q=nightSheet_(NIGHT.queue,H.queue);if(q.getMaxRows()<q.getLastRow()+tasks.length)q.insertRowsAfter(q.getMaxRows(),q.getLastRow()+tasks.length-q.getMaxRows());q.getRange(q.getLastRow()+1,1,tasks.length,16).setValues(tasks);const ctl=nightSheet_(NIGHT.control,H.control);ctl.appendRow([id,nightNow_(),nightNow_(),nightNow_(),tasks.length,0,0,0,0,'진행중','','무료 할당량 내 분할 실행 / 중복 검사 없음']);props.setProperties({NIGHT_ACTIVE:'1',NIGHT_RUN:id});nightInstallTick_();return id+' : '+tasks.length+'건 등록 완료';});}
function nightInstallTick_(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT.tickFn).forEach(t=>ScriptApp.deleteTrigger(t));ScriptApp.newTrigger(NIGHT.tickFn).timeBased().everyMinutes(NIGHT.tickMinutes).create();}
/** 기존 작업대기열 그대로 재개: 신규 등록/중복검사/결과 초기화 없음. */
function nightResume(){
  return nightLock_(()=>{
    nightValidate_();
    const props=PropertiesService.getScriptProperties();
    if(props.getProperty('NIGHT_ACTIVE')==='1')return '이미 실행 중';
    const id=props.getProperty('NIGHT_RUN');
    if(!id)throw Error('재개할 기존 실행ID가 없습니다.');
    const q=nightSheet_(NIGHT.queue,H.queue);
    const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues():[];
    let waiting=0;
    for(let i=0;i<rows.length;i++){
      const r=rows[i];
      if(r[0]!==id)continue;
      if(r[8]==='진행중')q.getRange(i+2,8,1,2).setValues([['대기','대기']]);
      if(r[8]==='재시도'){
        q.getRange(i+2,10,1,3).setValues([[0,'','']]);
      }
      if(['대기','재시도','진행중'].includes(r[8]))waiting++;
    }
    if(!waiting)throw Error('재개할 대기 작업이 없습니다.');
    props.setProperty('NIGHT_ACTIVE','1');
    nightInstallTick_();
    nightUpdateControl_(id);
    return '기존 실행 '+id+' / 남은 작업 '+waiting+'건 재개';
  });
}
/** 10자리 검증 코드 오류로 보류된 행만 기존 결과행을 유지하며 재처리 */
function nightRepairBadBizNo(){
 return nightLock_(()=>{
  const id=PropertiesService.getScriptProperties().getProperty('NIGHT_RUN');
  if(!id)throw Error('기존 실행ID가 없습니다.');
  const q=nightSheet_(NIGHT.queue,H.queue),result=nightSheet_(NIGHT.result,H.result);
  const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues():[];
  let count=0;
  for(let i=0;i<rows.length;i++){
   const r=rows[i],outRow=Number(r[14]);
   if(r[0]!==id||r[8]!=='보류'||!Number.isInteger(outRow)||outRow<2||outRow>result.getLastRow())continue;
   if(String(result.getRange(outRow,12).getDisplayValue())!=='사업자번호 10자리 확인 필요')continue;
   q.getRange(i+2,8,1,5).setValues([['대기','재시도',0,'','']]);
   count++;
  }
  nightSheet_('야간분석_진단기록',['실행시각','진단함수','결과','소요초','검사내용']).appendRow([nightNow_(),'nightRepairBadBizNo','성공',0,'오류 보류 재처리 '+count+'건']);
  return '오류 보류 '+count+'건 재처리 대기';
 });
}
/** 자동 감시 설치: 최초 1회만 수동 실행. 기존 TOP10 및 다른 트리거는 건드리지 않음. */
function nightAutoSetup(){
 return nightLock_(()=>{
  nightValidate_();
  const fn='nightWatchdog';
  const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===fn);
  if(!triggers.length)ScriptApp.newTrigger(fn).timeBased().everyMinutes(10).create();
  const p=PropertiesService.getScriptProperties(),id=p.getProperty('NIGHT_RUN');
  const q=nightSheet_(NIGHT.queue,H.queue);
  const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues().filter(r=>r[0]===id):[];
  const pending=rows.filter(r=>['대기','재시도','진행중'].includes(r[8])).length;
  if(pending){
   p.setProperty('NIGHT_ACTIVE','1');
   if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()===NIGHT.tickFn))nightInstallTick_();
  }
  nightHealthWrite_('설치완료','감시 10분 간격 / 남은 작업 '+pending+'건');
  return '자동감시 설치 완료 / 남은 작업 '+pending+'건';
 });
}
function nightHealthWrite_(state,detail){
 const p=PropertiesService.getScriptProperties();
 const triggers=ScriptApp.getProjectTriggers();
 const id=p.getProperty('NIGHT_RUN')||'';
 const q=nightSheet_(NIGHT.queue,H.queue);
 const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues().filter(r=>r[0]===id):[];
 const count=x=>rows.filter(r=>r[8]===x).length;
 const health=nightSheet_('야간분석_자동진단',['진단시각','실행ID','활성상태','처리트리거','감시트리거','완료','보류','대기','재시도','실패','진단상태','상세']);
 health.appendRow([nightNow_(),id,p.getProperty('NIGHT_ACTIVE')||'0',triggers.filter(t=>t.getHandlerFunction()===NIGHT.tickFn).length,triggers.filter(t=>t.getHandlerFunction()==='nightWatchdog').length,count('완료'),count('보류'),count('대기'),count('재시도'),count('실패'),state,String(detail||'').slice(0,400)]);
}
/** 10분마다 상태를 시트에 남기고, 트리거 누락/실행 정체를 자동 복구. */
function nightWatchdog(){
 try{
  const p=PropertiesService.getScriptProperties(),id=p.getProperty('NIGHT_RUN');
  if(!id){nightHealthWrite_('대기','실행ID 없음');return;}
  const q=nightSheet_(NIGHT.queue,H.queue);
  const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues().filter(r=>r[0]===id):[];
  const pending=rows.some(r=>['대기','재시도','진행중'].includes(r[8]));
  if(p.getProperty('NIGHT_ACTIVE')!=='1'){nightHealthWrite_('수동중지','nightResume 실행 전 자동재개 금지');return;}
  if(!pending){nightHealthWrite_('작업종료','남은 작업 없음');return;}
  const ticks=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT.tickFn);
  if(!ticks.length){nightInstallTick_();nightHealthWrite_('트리거복구','처리 트리거 재설치');}
  if(p.getProperty('NIGHT_ACTIVE')!=='1'){p.setProperty('NIGHT_ACTIVE','1');nightHealthWrite_('실행복구','활성 상태 복원');}
  const ctl=nightBook_().getSheetByName(NIGHT.control);
  let updated=0;
  if(ctl){
   const data=ctl.getDataRange().getValues();
   const line=data.find((r,i)=>i>0&&r[0]===id);
   if(line&&line[3] instanceof Date)updated=line[3].getTime();
  }
  const lastStart=Number(p.getProperty('NIGHT_TICK_STARTED')||0);
  const recentlyStarted=lastStart>0&&Date.now()-lastStart<9*60*1000;
  if(!recentlyStarted&&(!updated||Date.now()-updated>12*60*1000)){
   nightHealthWrite_('정체복구시도','12분 이상 처리기록 없음: 직접 실행');
   nightTick();
  }else nightHealthWrite_('정상감시',recentlyStarted?'처리 함수 실행 중 또는 최근 실행':'자동 처리 기록 확인');
 }catch(e){
  try{nightHealthWrite_('감시오류',String(e.message||e));}catch(ignore){}
  console.error('nightWatchdog: '+String(e.stack||e));
 }
}
/** 기존 시험 실행에서 기업명 중복으로 보류된 항목만 재검증. 기존 결과행 유지. */
/** 비공시 기업으로 확인된 기존 보류 행만 재분석 예약. 결과 행 덮어쓰기, 새 행 중복 추가 방지. */
function nightRecheckNonDart(){
 return nightLock_(()=>{
  const p=PropertiesService.getScriptProperties(),id=p.getProperty('NIGHT_RUN');
  if(!id)throw Error('기존 실행ID 없음');
  if(p.getProperty('NIGHT_ACTIVE')==='1')throw Error('먼저 nightStop으로 분석을 중지하세요.');
  const q=nightSheet_(NIGHT.queue,H.queue),out=nightSheet_(NIGHT.result,H.result);
  const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues():[];
  let count=0;
  for(let i=0;i<rows.length;i++){
   const r=rows[i],outRow=Number(r[14]);
   if(r[0]!==id||r[8]!=='보류'||!Number.isInteger(outRow)||outRow<2||outRow>out.getLastRow())continue;
   if(String(out.getRange(outRow,12).getDisplayValue())!=='DART 기업명 일치 항목 없음')continue;
   q.getRange(i+2,8,1,5).setValues([['대기','재시도',0,'','']]);
   count++;
  }
  return 'DART 미등록 기존 보류 '+count+'건 재분석 예약. nightResume으로 재개하세요.';
 });
}
function nightRetryAmbiguous(){
 return nightLock_(()=>{
  const p=PropertiesService.getScriptProperties(),id=p.getProperty('NIGHT_RUN');
  if(!id)throw Error('기존 실행ID 없음');
  if(p.getProperty('NIGHT_ACTIVE')==='1')throw Error('현재 분석 진행 중');
  const q=nightSheet_(NIGHT.queue,H.queue),out=nightSheet_(NIGHT.result,H.result);
  const rows=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues():[];
  let count=0;
  for(let i=0;i<rows.length;i++){
   const r=rows[i],resultRow=Number(r[14]);
   if(r[0]!==id||r[8]!=='보류'||!Number.isInteger(resultRow)||resultRow<2||resultRow>out.getLastRow())continue;
   const reason=String(out.getRange(resultRow,12).getDisplayValue());
   if(reason!=='동명 또는 유사 기업 복수: DART 고유번호 수동 확인 필요')continue;
   q.getRange(i+2,8,1,5).setValues([['대기','재시도',0,'','']]);
   count++;
  }
  if(count){
   p.setProperty('NIGHT_ACTIVE','1');
   nightInstallTick_();
   nightUpdateControl_(id);
  }
  nightSheet_('야간분석_진단기록',['실행시각','진단함수','결과','소요초','검사내용']).appendRow([nightNow_(),'nightRetryAmbiguous','성공',0,'동명 기업 '+count+'건 재검증 예약']);
  return '동명 기업 '+count+'건 재검증 예약';
 });
}
function nightStop(){nightLock_(()=>{PropertiesService.getScriptProperties().setProperty('NIGHT_ACTIVE','0');nightClearTick_();});}
function nightClearTick_(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT.tickFn).forEach(t=>ScriptApp.deleteTrigger(t));}
/** Google 서버 트리거: 실패 행은 기록하고 다음 회사로 진행합니다. */
function nightTick(){if(PropertiesService.getScriptProperties().getProperty('NIGHT_ACTIVE')!=='1')return;PropertiesService.getScriptProperties().setProperty('NIGHT_TICK_STARTED',String(Date.now()));nightLock_(()=>{const start=Date.now(),id=PropertiesService.getScriptProperties().getProperty('NIGHT_RUN');const q=nightSheet_(NIGHT.queue,H.queue);const n=q.getLastRow();if(n<2){PropertiesService.getScriptProperties().setProperty('NIGHT_ACTIVE','0');nightClearTick_();return;}const rows=q.getRange(2,1,n-1,16).getValues();let done=0;for(let i=0;i<rows.length&&done<NIGHT.maxPerTick&&Date.now()-start<NIGHT.maxRuntimeMs;i++){const r=rows[i];if(r[0]!==id||!['대기','재시도'].includes(r[8]))continue;if(r[10]&&new Date(r[10]).getTime()>Date.now())continue;const row=i+2;done++;try{q.getRange(row,8,1,2).setValues([['DART 분석','진행중']]);const outcome=nightAnalyze_(r);const result=nightSheet_(NIGHT.result,H.result);const oldResultRow=Number(r[14]);const reuse=Number.isInteger(oldResultRow)&&oldResultRow>=2&&oldResultRow<=result.getLastRow();if(reuse)result.getRange(oldResultRow,1,1,H.result.length).setValues([outcome.result]);else result.appendRow(outcome.result);const resultRow=reuse?oldResultRow:result.getLastRow();const agent=nightSheet_(NIGHT.agent,H.agent);const agentValues=agent.getLastRow()>1?agent.getRange(2,1,agent.getLastRow()-1,4).getDisplayValues():[];const agentIndex=reuse?agentValues.findIndex(a=>a[0]===String(id)&&String(a[3]).replace(/\D/g,'')===String(r[3]).replace(/\D/g,'')):-1;const agentRow=[id,'',r[2],r[3],r[4],r[5],r[6],outcome.signal,outcome.level,outcome.evidence,'검토대기','배정대기','재무징후는 실제 미수채권 보유의 증거가 아님'];if(agentIndex>=0)agent.getRange(agentIndex+2,1,1,H.agent.length).setValues([agentRow]);else agent.appendRow(agentRow);q.getRange(row,8,1,8).setValues([['완료',outcome.status,Number(r[9])+1,'','',outcome.corpCode,nightNow_(),resultRow]]);}catch(e){const tries=Number(r[9])+1;const retry=tries<NIGHT.maxRetries;const delay=Math.min(60,Math.pow(2,tries)*5);const next=retry?new Date(Date.now()+delay*60000):'';q.getRange(row,8,1,5).setValues([['DART 분석',retry?'재시도':'실패',tries,next,String(e.message||e).slice(0,500)]]);nightSheet_(NIGHT.errors,['실행ID','기업명','사업자등록번호','단계','오류유형','오류내용','발생시각','재시도결과','최종처리']).appendRow([id,r[2],r[3],'DART',retry?'재시도':'최종실패',String(e.message||e),nightNow_(),retry?'예정':'없음',retry?'대기':'실패']);}}
nightUpdateControl_(id);});}
function nightUpdateControl_(id){const q=nightSheet_(NIGHT.queue,H.queue);const values=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues().filter(r=>r[0]===id):[];const count=s=>values.filter(r=>r[8]===s).length;const completed=count('완료'),partial=count('부분완료'),failed=count('실패'),held=count('보류');const waiting=values.length-completed-partial-failed-held;const ctl=nightSheet_(NIGHT.control,H.control);const c=ctl.getDataRange().getValues();for(let i=1;i<c.length;i++)if(c[i][0]===id){ctl.getRange(i+1,4,1,9).setValues([[nightNow_(),values.length,completed,partial,held,failed,waiting?'진행중':'완료',waiting?'약 5분 간격으로 재개':'',waiting?'API 할당량/오류에 따라 처리 시간 변동':'완료: 결과 검토 필요']]);break;}if(!waiting){PropertiesService.getScriptProperties().setProperty('NIGHT_ACTIVE','0');nightClearTick_();}}
function nightAnalyze_(r){const name=String(r[2]).trim();const reg=String(r[3]||'').replace(/\D/g,'');const phone=String(r[4]||'');const address=String(r[5]||'');const found=nightCorp_(name);const date=nightDate_(nightNow_());if(found.length>1){const verifiedMatches=[];for(const candidate of found.slice(0,12)){const check=nightVerifyCorp_(candidate.code,reg);if(check.ok)verifiedMatches.push(candidate);}if(verifiedMatches.length===1){found.splice(0,found.length,verifiedMatches[0]);}else if(verifiedMatches.length>1){found.splice(0,found.length,...verifiedMatches);}}if(found.length!==1){const reason=found.length?'동명 또는 유사 기업 복수: DART 고유번호 수동 확인 필요':'DART 기업명 일치 항목 없음';if(!found.length){const signal='공시재무 없음 / 업종 기반 거래구조 검토';const evidence='비즈라인 제공 기업명·사업자번호·주소·업종 / DART 일치 기업 없음 / 국세청 상태 미검증';return {corpCode:'',status:'부분완료',signal,level:'비공시·기초정보',evidence,result:[name,reg,phone,address,'','비즈라인 원본 사업자번호 기재 / 국세청 미검증','','','공시재무 및 연체·충당금 확인 불가','부분완료','공시재무 미확인: 업종별 거래구조·외상거래 가능성은 담당자 확인 필요','DART 기업명 일치 항목 없음 (기업 부존재를 뜻하지 않음)',date+' / '+evidence]};}return {corpCode:'',status:'보류',signal:'미검증',level:'보류',evidence:reason,result:[name,reg,phone,address,'','사업자번호 국세청 미검증','','','', '보류','재무 분석 불가',reason,date+' / DART 기업개황']};}
const corp=found[0];
const verified=nightVerifyCorp_(corp.code,reg);
if(!verified.ok){
  const reason=verified.reason;
  return {corpCode:corp.code,status:'보류',signal:'미검증',level:'보류',evidence:reason,result:[name,reg,phone,address,corp.code,'DART 사업자번호 대조 불일치 또는 미확인','','','','보류','재무 분석 보류',reason,date+' / DART 기업개황']};
}
const y25=nightFinance_(corp.code,2025);const y24=nightFinance_(corp.code,2024);const basis=y25.basis&&y24.basis&&y25.basis===y24.basis?y25.basis:(y25.basis||y24.basis||'');const comparable=!!y25.basis&&y25.basis===y24.basis;const fmt=(f,y)=>y+' 매출 '+nightWon_(f.revenue)+', 매출채권 '+nightWon_(f.receivable);const changes=fmt(y24,2024)+' / '+fmt(y25,2025);const enough=comparable&&y25.revenue!==null&&y24.revenue!==null&&y25.receivable!==null&&y24.receivable!==null;const status=enough?'완료':'부분완료';const signal=!comparable?'연도별 재무제표 기준 불일치 또는 누락':!enough?'재무항목 누락: 추가 확인 필요':y25.receivable>y24.receivable?'매출채권 증가 (연체 여부 미확인)':'재무검토 필요';const note='연체·손실충당금 주석 미검증: 숫자 추정 금지';const evidence='DART '+corp.code+' / '+basis+' / 2024·2025 사업보고서 재무제표 / 금액 원';return {corpCode:corp.code,status,signal,level:status,evidence,result:[name,reg,phone,address,corp.code,'DART 기업명 및 사업자번호 일치 / 국세청 영업상태 미검증',basis,changes,note,status,signal,status==='완료'?'':(!comparable?'2024·2025 연결/별도 기준 불일치 또는 누락':'연도별 매출·매출채권 항목 일부 누락'),date+' / '+evidence]};}
function nightWon_(n){return n===null?'확인 불가':Math.round(n).toLocaleString('ko-KR')+'원';}
function nightGet_(url){const response=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true});const code=response.getResponseCode();if(code===429||code>=500)throw Error('API 일시 오류 HTTP '+code);if(code!==200)throw Error('API HTTP '+code);return response;}
/** DART 고유번호 전체 목록을 읽어 기업명 '정확히 일치'만 선택; 유사명 자동매칭 금지. */
/**
 * DART 기업목록: 한 번 내려받은 ZIP을 Drive 비공개 파일에 24시간 보관합니다.
 * 매 틱마다 API를 호출하지 않으며, 틱 내에서도 압축 해제와 XML 스캔을 한 번만 합니다.
 * 목록은 정확한 기업명으로만 매칭합니다. 동명이인은 자동으로 선택하지 않습니다.
 * 주의: ZIP 파일은 해당 Google 계정의 Drive에 저장되며 공유하지 않습니다.
 */
var NIGHT_CORP_INDEX = null;
var NIGHT_CORP_NORMALIZED=null;
function nightCorpIndex_(){
  if(NIGHT_CORP_INDEX !== null)return NIGHT_CORP_INDEX;
  const props=PropertiesService.getScriptProperties();
  const stamp=Number(props.getProperty('NIGHT_CORP_CACHE_TIME')||0);
  const fileId=props.getProperty('NIGHT_CORP_CACHE_FILE');
  let zipped=null;
  if(fileId && Date.now()-stamp<7*24*60*60*1000){
    try{zipped=DriveApp.getFileById(fileId).getBlob();}catch(e){zipped=null;}
  }
  if(!zipped){
    const url='https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key='+encodeURIComponent(nightKey_());
    zipped=nightGet_(url).getBlob();
    // ZIP은 다운로드 직후 반복 해제하지 않고 캐시 저장 후 한 번만 해제합니다.
    // 다운로드 정상 여부와 압축 해제 오류를 별도로 기록합니다.
    const raw=zipped.getBytes();
    if(raw.length<4 || raw[0]!==80 || raw[1]!==75){
      throw Error('DART 기업목록 API 응답이 ZIP이 아닙니다 (HTTP 200, 바이트 '+raw.length+')');
    }
    const oldId=props.getProperty('NIGHT_CORP_CACHE_FILE');
    const cache=DriveApp.createFile(zipped.setName('night-dart-corpcode-cache.zip'));
    props.setProperties({NIGHT_CORP_CACHE_FILE:cache.getId(),NIGHT_CORP_CACHE_TIME:String(Date.now())});
    if(oldId){try{DriveApp.getFileById(oldId).setTrashed(true);}catch(e){}}
  }
  // DART responds with application/x-msdownload; Apps Script unzip requires application/zip.
  zipped=zipped.setContentType('application/zip');
  let xml;
  try{xml=Utilities.unzip(zipped)[0].getDataAsString('UTF-8');}
  catch(e){props.deleteProperty('NIGHT_CORP_CACHE_TIME');throw Error('DART ZIP 압축 해제 실패: '+String(e.message||e).slice(0,180));}
  const index=Object.create(null);
  const blocks=xml.match(/<list>[\s\S]*?<\/list>/g)||[];
  if(!blocks.length)throw Error('DART 기업목록 데이터가 비어 있습니다.');
  for(const block of blocks){
    const nm=/<corp_name>([\s\S]*?)<\/corp_name>/.exec(block);
    const code=/<corp_code>([^<]+)<\/corp_code>/.exec(block);
    if(!nm||!code)continue;
    const name=nm[1].replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').trim();
    if(!index[name])index[name]=[];
    index[name].push({code:code[1]});
  }
  // 기업목록 1회 해제 시 정규화된 이름별 색인도 구축하여 반복 전체 검색을 제거합니다.
  const normalized=Object.create(null);
  for(const nameInDart in index){
    const key=nightCorpName_(nameInDart);
    if(!key)continue;
    if(!normalized[key])normalized[key]=[];
    normalized[key].push(...index[nameInDart]);
  }
  NIGHT_CORP_NORMALIZED=normalized;
  NIGHT_CORP_INDEX=index;
  return index;
}
function nightCorpName_(name){
  return String(name||'').replaceAll('주식회사','').replaceAll('(주)','').replaceAll('㈜','').replaceAll(' ','').trim();
}
function nightCorp_(name){
  const index=nightCorpIndex_();
  const exact=index[String(name).trim()]||[];
  if(exact.length)return exact;
  const key=nightCorpName_(name);
  if(!key)return [];
  return (NIGHT_CORP_NORMALIZED&&NIGHT_CORP_NORMALIZED[key])||[];
}
function nightVerifyCorp_(corpCode,bizNo){
  if(!/^\d{10}$/.test(bizNo))return {ok:false,reason:'사업자번호 10자리 확인 필요'};
  const url='https://opendart.fss.or.kr/api/company.json?crtfc_key='+encodeURIComponent(nightKey_())+'&corp_code='+encodeURIComponent(corpCode);
  const data=JSON.parse(nightGet_(url).getContentText('UTF-8'));
  if(data.status!=='000')return {ok:false,reason:'DART 기업개황 조회 실패: '+String(data.status||'응답 없음')};
  const dartNo=String(data.bizr_no||'').replace(/\D/g,'');
  if(!dartNo)return {ok:false,reason:'DART 기업개황에 사업자번호 없음'};
  return dartNo===bizNo?{ok:true,reason:'사업자번호 일치'}:{ok:false,reason:'입력 사업자번호와 DART 사업자번호 불일치'};
}
function nightFinance_(corp,year){const key=encodeURIComponent(nightKey_());for(const basis of ['CFS','OFS']){const url='https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json?crtfc_key='+key+'&corp_code='+encodeURIComponent(corp)+'&bsns_year='+year+'&reprt_code=11011&fs_div='+basis;const data=JSON.parse(nightGet_(url).getContentText('UTF-8'));if(data.status==='013')continue;if(data.status!=='000')throw Error('DART 재무 '+year+' '+basis+': '+(data.message||data.status));const list=data.list||[];let revenue=null,receivable=null;for(const x of list){const nm=String(x.account_nm||'').replace(/\s/g,'');const amt=nightAmount_(x.thstrm_amount);if(amt===null)continue;if(revenue===null&&/^(매출액|수익\(매출액\)|영업수익)$/.test(nm))revenue=amt;if(receivable===null&&/^(매출채권|매출채권및기타채권)$/.test(nm))receivable=amt;}if(revenue!==null||receivable!==null)return {basis,revenue,receivable};}return {basis:'',revenue:null,receivable:null};}
function nightAmount_(v){if(v===null||v===undefined||String(v).trim()==='-'||String(v).trim()==='')return null;const n=Number(String(v).replace(/,/g,''));return Number.isFinite(n)?n:null;}
/** 설치 후 확인용. 인증정보 자체는 표시하지 않습니다. */
function nightStatus(){return {dartKeyRegistered:!!nightKey_(),active:PropertiesService.getScriptProperties().getProperty('NIGHT_ACTIVE')==='1',run:PropertiesService.getScriptProperties().getProperty('NIGHT_RUN')||'',tickTriggers:ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT.tickFn).length};}


/** DART ZIP response diagnostic; never prints API key or response body. */
function nightDartDiagnostic(){
  nightValidate_();
  const url='https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key='+encodeURIComponent(nightKey_().trim());
  const r=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
  const bytes=r.getBlob().getBytes();
  const zip=bytes.length>3&&bytes[0]===80&&bytes[1]===75;
  console.log(JSON.stringify({http:r.getResponseCode(),zip:zip,bytes:bytes.length,contentType:r.getHeaders()['Content-Type']||'unknown'}));
}

/** 독립 시험: DART 기업목록 ZIP 압축 해제 및 지정 기업 검색만 확인. 작업대기열은 변경하지 않음. */
function nightCorpTest(){
  const started=Date.now();
  const names=['티케이지휴켐스(주)','(주)에이블씨엔씨','한전산업개발(주)'];
  let status='성공', detail='';
  try{
    const index=nightCorpIndex_();
    const result=names.map(name=>({name:name,exactMatches:(index[name]||[]).length,normalizedMatches:nightCorp_(name).length}));
    detail=JSON.stringify(result);
  }catch(e){
    status='오류';
    detail=String(e.message||e).slice(0,400);
  }
  const seconds=Math.round((Date.now()-started)/1000);
  const log=nightSheet_('야간분석_진단기록',['실행시각','진단함수','결과','소요초','검사내용']);
  log.appendRow([nightNow_(),'nightCorpTest',status,seconds,detail]);
  console.log(JSON.stringify({status:status,elapsedSeconds:seconds,detail:detail}));
}


/**
 * 국세청 사업자등록 상태조회 (공공데이터포털 NTS API).
 * 대량분석 200건의 원본 사업자번호를 검증합니다. DART 결과는 변경하지 않습니다.
 * 별도 스크립트 속성 NTS_API_KEY 필요. 인증키를 코드/시트/로그에 기록하지 않습니다.
 * 실행: nightNtsStart() 1회 -> 자동 트리거 nightNtsTick() (5분마다 100건).
 */
const NIGHT_NTS = Object.freeze({
 sheet:'대량분석_국세청검증', fn:'nightNtsTick', size:100,
 headers:['실행ID','원본행','기업명','사업자등록번호','조회상태','국세청 사업자상태','과세유형','폐업일자','국세청 상태코드','조회시각','오류·참고']
});
function nightNtsKey_(){
 const p=PropertiesService.getScriptProperties();
 return String(p.getProperty('NTS_API_KEY')||p.getProperty('NTS_SERVICE_KEY')||'').trim();
}
function nightNtsInstall_(){
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT_NTS.fn).forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger(NIGHT_NTS.fn).timeBased().everyMinutes(5).create();
}
function nightNtsStop(){
 PropertiesService.getScriptProperties().setProperty('NIGHT_NTS_ACTIVE','0');
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT_NTS.fn).forEach(t=>ScriptApp.deleteTrigger(t));
 return '국세청 상태조회 중지';
}
function nightNtsStart(){
 return nightLock_(()=>{
  if(!nightNtsKey_())throw Error('이 Apps Script 프로젝트의 스크립트 속성 NTS_API_KEY가 없습니다. 인증키는 채팅이나 코드에 입력하지 마세요.');
  const p=PropertiesService.getScriptProperties(),run=p.getProperty('NIGHT_RUN');
  if(!run)throw Error('기존 대량분석 실행ID 없음');
  if(p.getProperty('NIGHT_NTS_ACTIVE')==='1')return '국세청 조회가 이미 실행 중입니다.';
  const queue=nightSheet_(NIGHT.queue,H.queue),rows=queue.getLastRow()>1?queue.getRange(2,1,queue.getLastRow()-1,16).getDisplayValues():[];
  const target=rows.map(r=>({r})).filter(x=>x.r[0]===run);
  if(!target.length)throw Error('현재 대량분석 실행ID에 해당하는 기업 없음');
  const sheet=nightSheet_(NIGHT_NTS.sheet,NIGHT_NTS.headers);
  const old=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,4).getDisplayValues():[];
  const exists=new Set(old.map(r=>r[0]+'|'+r[1]));
  const append=[];
  for(const x of target){
   const r=x.r,key=run+'|'+r[1];
   if(exists.has(key))continue;
   append.push([run,r[1],r[2],String(r[3]||'').replace(/\D/g,''),'대기','','','','','','']);
  }
  if(append.length){
   if(sheet.getMaxRows()<sheet.getLastRow()+append.length)sheet.insertRowsAfter(sheet.getMaxRows(),sheet.getLastRow()+append.length-sheet.getMaxRows());
   sheet.getRange(sheet.getLastRow()+1,1,append.length,NIGHT_NTS.headers.length).setValues(append);
  }
  p.setProperties({NIGHT_NTS_ACTIVE:'1',NIGHT_NTS_RUN:run});
  nightNtsInstall_();
  return '국세청 상태조회 '+run+' / 신규 '+append.length+'건 예약. 결과는 '+NIGHT_NTS.sheet+' 시트';
 });
}
function nightNtsTick(){
 if(PropertiesService.getScriptProperties().getProperty('NIGHT_NTS_ACTIVE')!=='1')return;
 return nightLock_(()=>{
  const p=PropertiesService.getScriptProperties(),run=p.getProperty('NIGHT_NTS_RUN');
  const sheet=nightSheet_(NIGHT_NTS.sheet,NIGHT_NTS.headers);
  const values=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,NIGHT_NTS.headers.length).getDisplayValues():[];
  const todo=[];
  for(let i=0;i<values.length&&todo.length<NIGHT_NTS.size;i++){
   const r=values[i];
   if(r[0]===run&&['대기','재시도'].includes(r[4]))todo.push({row:i+2,biz:String(r[3]||'').replace(/\D/g,'')});
  }
  if(!todo.length){nightNtsStop();return '국세청 상태조회 종료';}
  const valid=todo.filter(x=>/^\d{10}$/.test(x.biz));
  for(const x of todo.filter(x=>!/^\d{10}$/.test(x.biz))){
   sheet.getRange(x.row,5,1,7).setValues([['확인필요','','','','',nightNow_(),'사업자번호 10자리 아님']]);
  }
  if(!valid.length)return '사업자번호 오류 '+todo.length+'건 기록';
  const key=nightNtsKey_();
  if(!key){nightNtsStop();throw Error('NTS_API_KEY가 없습니다.');}
  const url='https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey='+encodeURIComponent(key);
  let response,parsed;
  try{
   response=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify({b_no:valid.map(x=>x.biz)}),muteHttpExceptions:true});
   if(response.getResponseCode()!==200)throw Error('국세청 API HTTP '+response.getResponseCode());
   parsed=JSON.parse(response.getContentText());
   if(!Array.isArray(parsed.data))throw Error('국세청 API 응답에 data 배열 없음');
  }catch(e){
   // 키와 전체 API 응답을 로그에 기록하지 않습니다.
   for(const x of valid)sheet.getRange(x.row,5,1,7).setValues([['재시도','','','','',nightNow_(),String(e.message||e).slice(0,180)]]);
   return '국세청 API 일시 오류: '+valid.length+'건 다음 실행 재시도';
  }
  const byNo=new Map(parsed.data.map(v=>[String(v.b_no||'').replace(/\D/g,''),v]));
  for(const x of valid){
   const v=byNo.get(x.biz);
   if(!v){sheet.getRange(x.row,5,1,7).setValues([['확인필요','','','','',nightNow_(),'국세청 응답에 해당 사업자번호 없음']]);continue;}
   const code=String(v.b_stt_cd||'');
   const label=String(v.b_stt||'');
   const state=code==='01'?'계속사업':code==='02'?'휴업':code==='03'?'폐업':'확인필요';
   sheet.getRange(x.row,5,1,7).setValues([['조회완료',state+' / '+label,String(v.tax_type||''),String(v.end_dt||''),code,nightNow_(),state==='확인필요'?'상태코드 미확인':'국세청 API 사업자등록 상태조회']]);
  }
  const latest=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,5).getDisplayValues():[];
  const remaining=latest.some(r=>r[0]===run&&['대기','재시도'].includes(r[4]));
  if(!remaining&&todo.length<NIGHT_NTS.size)nightNtsStop();
  return '국세청 '+valid.length+'건 조회 완료';
 });
}
