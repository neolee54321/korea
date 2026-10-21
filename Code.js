/**
 * 고려신용정보 | 독립형 무료 야간분석 v1.0 (2026-10-21)
 * Google Apps Script 신규 '독립형 프로젝트'에만 설치하세요.
 * 기존 TOP10 스크립트, 트리거, 시트는 수정하지 않습니다.
 * 무료 할당량 내 분할 실행. 실제 미수채권 존재/외부위탁 가능성은 추정하지 않습니다.
 * DART OpenAPI 사용: https://opendart.fss.or.kr/
 */
const NIGHT = Object.freeze({
  ssid:'1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU',
  input:'신규20_입력대기', queue:'야간분석_작업대기열',
  control:'야간분석_실행관리', errors:'야간분석_오류기록',
  result:'신규20_시험결과', agent:'관리사배분_결과',
  tickFn:'nightTick', tickMinutes:5, maxPerTick:5,
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
function nightStart(){return nightLock_(()=>{nightValidate_();const props=PropertiesService.getScriptProperties();if(props.getProperty('NIGHT_ACTIVE')==='1')throw Error('이미 야간분석이 진행 중입니다.');const input=nightBook_().getSheetByName(NIGHT.input);if(!input)throw Error('입력 시트가 없습니다.');const values=input.getDataRange().getDisplayValues().slice(1);const tasks=[];const id=Utilities.formatDate(new Date(),'Asia/Seoul','yyyyMMdd-HHmmss');values.forEach((v,i)=>{if(!String(v[1]||'').trim())return;tasks.push([id,i+2,v[1],v[3],v[2],v[5],v[6],'대기','대기',0,'','','','','','비즈라인']);});if(!tasks.length)throw Error('입력 기업이 없습니다.');const q=nightSheet_(NIGHT.queue,H.queue);if(q.getMaxRows()<q.getLastRow()+tasks.length)q.insertRowsAfter(q.getMaxRows(),q.getLastRow()+tasks.length-q.getMaxRows());q.getRange(q.getLastRow()+1,1,tasks.length,16).setValues(tasks);const ctl=nightSheet_(NIGHT.control,H.control);ctl.appendRow([id,nightNow_(),nightNow_(),nightNow_(),tasks.length,0,0,0,0,'진행중','','무료 할당량 내 분할 실행 / 중복 검사 없음']);props.setProperties({NIGHT_ACTIVE:'1',NIGHT_RUN:id});nightInstallTick_();return id+' : '+tasks.length+'건 등록 완료';});}
function nightInstallTick_(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT.tickFn).forEach(t=>ScriptApp.deleteTrigger(t));ScriptApp.newTrigger(NIGHT.tickFn).timeBased().everyMinutes(NIGHT.tickMinutes).create();}
function nightStop(){nightLock_(()=>{PropertiesService.getScriptProperties().setProperty('NIGHT_ACTIVE','0');nightClearTick_();});}
function nightClearTick_(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===NIGHT.tickFn).forEach(t=>ScriptApp.deleteTrigger(t));}
/** Google 서버 트리거: 실패 행은 기록하고 다음 회사로 진행합니다. */
function nightTick(){if(PropertiesService.getScriptProperties().getProperty('NIGHT_ACTIVE')!=='1')return;nightLock_(()=>{const start=Date.now(),id=PropertiesService.getScriptProperties().getProperty('NIGHT_RUN');const q=nightSheet_(NIGHT.queue,H.queue);const n=q.getLastRow();if(n<2){PropertiesService.getScriptProperties().setProperty('NIGHT_ACTIVE','0');nightClearTick_();return;}const rows=q.getRange(2,1,n-1,16).getValues();let done=0;for(let i=0;i<rows.length&&done<NIGHT.maxPerTick&&Date.now()-start<NIGHT.maxRuntimeMs;i++){const r=rows[i];if(r[0]!==id||!['대기','재시도'].includes(r[8]))continue;if(r[10]&&new Date(r[10]).getTime()>Date.now())continue;const row=i+2;done++;try{q.getRange(row,8,1,2).setValues([['DART 분석','진행중']]);const outcome=nightAnalyze_(r);const result=nightSheet_(NIGHT.result,H.result);result.appendRow(outcome.result);const resultRow=result.getLastRow();const agent=nightSheet_(NIGHT.agent,H.agent);agent.appendRow([id,'',r[2],r[3],r[4],r[5],r[6],outcome.signal,outcome.level,outcome.evidence,'검토대기','배정대기','재무징후는 실제 미수채권 보유의 증거가 아님']);q.getRange(row,8,1,8).setValues([['완료',outcome.status,Number(r[9])+1,'','',outcome.corpCode,nightNow_(),resultRow]]);}catch(e){const tries=Number(r[9])+1;const retry=tries<NIGHT.maxRetries;const delay=Math.min(60,Math.pow(2,tries)*5);const next=retry?new Date(Date.now()+delay*60000):'';q.getRange(row,8,1,5).setValues([['DART 분석',retry?'재시도':'실패',tries,next,String(e.message||e).slice(0,500)]]);nightSheet_(NIGHT.errors,['실행ID','기업명','사업자등록번호','단계','오류유형','오류내용','발생시각','재시도결과','최종처리']).appendRow([id,r[2],r[3],'DART',retry?'재시도':'최종실패',String(e.message||e),nightNow_(),retry?'예정':'없음',retry?'대기':'실패']);}}
nightUpdateControl_(id);});}
function nightUpdateControl_(id){const q=nightSheet_(NIGHT.queue,H.queue);const values=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getValues().filter(r=>r[0]===id):[];const count=s=>values.filter(r=>r[8]===s).length;const completed=count('완료'),partial=count('부분완료'),failed=count('실패'),held=count('보류');const waiting=values.length-completed-partial-failed-held;const ctl=nightSheet_(NIGHT.control,H.control);const c=ctl.getDataRange().getValues();for(let i=1;i<c.length;i++)if(c[i][0]===id){ctl.getRange(i+1,4,1,8).setValues([[nightNow_(),values.length,completed,partial,held,failed,waiting?'진행중':'완료',waiting?'약 5분 간격으로 재개':'',waiting?'API 할당량/오류에 따라 처리 시간 변동':'완료: 결과 검토 필요']]);break;}if(!waiting){PropertiesService.getScriptProperties().setProperty('NIGHT_ACTIVE','0');nightClearTick_();}}
function nightAnalyze_(r){const name=String(r[2]).trim();const reg=String(r[3]||'').replace(/\D/g,'');const phone=String(r[4]||'');const address=String(r[5]||'');const found=nightCorp_(name);const date=nightDate_(nightNow_());if(found.length!==1){const reason=found.length?'동명 또는 유사 기업 복수: DART 고유번호 수동 확인 필요':'DART 기업명 일치 항목 없음';return {corpCode:'',status:'보류',signal:'미검증',level:'보류',evidence:reason,result:[name,reg,phone,address,'','사업자번호 국세청 미검증','','','', '보류','재무 분석 불가',reason,date+' / DART 기업개황']};}
const corp=found[0];const y25=nightFinance_(corp.code,2025);const y24=nightFinance_(corp.code,2024);const basis=y25.basis||y24.basis||'';const fmt=(f,y)=>y+' 매출 '+nightWon_(f.revenue)+', 매출채권 '+nightWon_(f.receivable);const changes=fmt(y24,2024)+' / '+fmt(y25,2025);const enough=y25.revenue!==null&&y24.revenue!==null;const status=enough?'완료':'부분완료';const signal=enough&&y25.receivable!==null&&y24.receivable!==null&&y25.receivable>y24.receivable?'매출채권 증가 (연체 여부 미확인)':'재무검토 필요';const note='연체·손실충당금 주석 미검증: 숫자 추정 금지';const evidence='DART '+corp.code+' / '+basis+' / 2024·2025 사업보고서 재무제표 / 금액 원';return {corpCode:corp.code,status,signal,level:status,evidence,result:[name,reg,phone,address,corp.code,'DART 기업명 일치 / 사업자번호 국세청 미검증',basis,changes,note,status,signal,status==='완료'?'':'연도별 항목 일부 누락',date+' / '+evidence]};}
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
function nightCorpIndex_(){
  if(NIGHT_CORP_INDEX !== null)return NIGHT_CORP_INDEX;
  const props=PropertiesService.getScriptProperties();
  const stamp=Number(props.getProperty('NIGHT_CORP_CACHE_TIME')||0);
  const fileId=props.getProperty('NIGHT_CORP_CACHE_FILE');
  let zipped=null;
  if(fileId && Date.now()-stamp<24*60*60*1000){
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
  NIGHT_CORP_INDEX=index;
  return index;
}
function nightCorp_(name){
  return nightCorpIndex_()[String(name).trim()]||[];
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
  try{
    const index=nightCorpIndex_();
    const result=names.map(name=>({name:name,exactMatches:(index[name]||[]).length}));
    console.log(JSON.stringify({ok:true,elapsedSeconds:Math.round((Date.now()-started)/1000),sample:result}));
  }catch(e){
    console.log(JSON.stringify({ok:false,elapsedSeconds:Math.round((Date.now()-started)/1000),error:String(e.message||e).slice(0,250)}));
  }
}
