/**
 * 대량입력 일별 스냅샷 및 야간 통합 운영 v1.
 * 기존 Code.js를 수정하지 않음. 같은 Apps Script 프로젝트에 추가 설치.
 * 야간 20시 시작, 07:30 결과 확정 시도. 트리거 실행시각은 구글 정책상 다소 변동.
 * 국세청/DART는 기존 엔진 사용. 나라장터 전수조회는 별도 엔진으로 미연결이며 미검증 표시.
 */
const DAILY={
 input:'대량분석_신규입력', output:'일일_배포대기', archive:'일일_분석이력',
 log:'일일_운영기록', start:'dailyStart', finalize:'dailyFinalize',
 headers:['실행ID','입력일','원본행','기업명','사업자등록번호','국세청상태','DART분석상태','재무징후','나라장터검증','영업검토순위','판정근거','배포상태']
};
function dailyTab_(name,headers){
 const ss=SpreadsheetApp.openById(NIGHT.ssid);let s=ss.getSheetByName(name);
 if(!s)s=ss.insertSheet(name);
 if(headers){if(s.getMaxColumns()<headers.length)s.insertColumnsAfter(s.getMaxColumns(),headers.length-s.getMaxColumns());s.getRange(1,1,1,headers.length).setValues([headers]);s.setFrozenRows(1);}
 return s;
}
function dailyLog_(type,msg){dailyTab_(DAILY.log,['시각','단계','내용']).appendRow([new Date(),type,String(msg).slice(0,500)]);}
function dailySnapshot_(){
 const s=SpreadsheetApp.openById(NIGHT.ssid).getSheetByName(DAILY.input);
 if(!s)throw Error('대량분석_신규입력 없음');
 const header=s.getRange(1,1,1,7).getDisplayValues()[0];
 const expected=['기업유형','기업명','전화번호','사업자번호','대표자명','기업주소','산업명'];
 if(expected.some((v,i)=>v!==header[i]))throw Error('입력 열 순서 변경 감지: 시작 중단');
 const values=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,7).getDisplayValues():[];
 const rows=values.map((r,i)=>({row:i+2,data:r})).filter(x=>x.data[1].trim());
 if(!rows.length)throw Error('입력 기업 없음');
 const invalid=rows.filter(x=>!/^[0-9]{10}$/.test(String(x.data[3]).replace(/\D/g,'')));
 if(invalid.length)throw Error('사업자번호 형식 오류 '+invalid.length+'건, 예: '+invalid.slice(0,5).map(x=>x.row).join(','));
 const payload=JSON.stringify(rows.map(x=>x.data));
 const hash=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,payload)).slice(0,36);
 return {rows,hash};
}
function dailyInstall(){
 const existing=ScriptApp.getProjectTriggers();
 for(const fn of [DAILY.start,DAILY.finalize])existing.filter(t=>t.getHandlerFunction()===fn).forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger(DAILY.start).timeBased().atHour(20).nearMinute(0).everyDays(1).inTimezone('Asia/Seoul').create();
 ScriptApp.newTrigger(DAILY.finalize).timeBased().atHour(7).nearMinute(30).everyDays(1).inTimezone('Asia/Seoul').create();
 dailyTab_(DAILY.output,DAILY.headers);dailyTab_(DAILY.archive,DAILY.headers);
 dailyLog_('설치','20시 입력 스냅샷 / 07:30 결과 작성 / 입력 원본 변경 없음');
 return '일별 자동화 설치. 테스트는 dailyStart() 수동 실행';
}
/** 동일한 입력내용으로 매일 반복 실행하는 실수를 방지. 기존 실행 중이면 건드리지 않음. */
function dailyStart(){
 const p=PropertiesService.getScriptProperties();
 try{
  const snap=dailySnapshot_();
  const today=Utilities.formatDate(new Date(),'Asia/Seoul','yyyyMMdd');
  if(p.getProperty('DAILY_INPUT_HASH')===snap.hash){dailyLog_('시작건너뜀','이전 입력과 동일: 새로운 업로드 미확인');return '입력 내용 변경 없음: 실행하지 않음';}
  if(p.getProperty('NIGHT_ACTIVE')==='1'||p.getProperty('NIGHT_NTS_ACTIVE')==='1')throw Error('기존 분석 진행 중: 새 실행 시작하지 않음');
  const run=nightStart(); // 이 함수가 입력을 대기열에 복제하므로 이후 입력 변경이 기존 실행에 영향을 주지 않음
  p.setProperties({DAILY_INPUT_HASH:snap.hash,DAILY_RUN:p.getProperty('NIGHT_RUN'),DAILY_DATE:today,DAILY_SIZE:String(snap.rows.length),DAILY_READY:'0'});
  try{nightNtsStart();}catch(e){dailyLog_('국세청시작오류',e.message||e);}
  try{dailyG2BStart();}catch(e){dailyLog_('나라장터시작오류',String(e.message||e));}
  dailyLog_('시작','입력 '+snap.rows.length+'건 / '+run);
  return run;
 }catch(e){dailyLog_('시작오류',e.message||e);throw e;}
}
/** 당일 실행ID에 속한 행만 조인. 이전 실행의 분석결과는 절대 섞지 않음. */
function dailyFinalize(){
 const p=PropertiesService.getScriptProperties(),run=p.getProperty('DAILY_RUN');
 if(!run)return '일별 실행ID 없음';
 const ss=SpreadsheetApp.openById(NIGHT.ssid);
 const q=ss.getSheetByName(NIGHT.queue),a=ss.getSheetByName(NIGHT.agent),n=ss.getSheetByName(NIGHT_NTS.sheet);
 const queue=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getDisplayValues().filter(r=>r[0]===run):[];
 if(!queue.length)throw Error('해당 실행ID의 대기열 없음');
 const agents=a&&a.getLastRow()>1?a.getRange(2,1,a.getLastRow()-1,13).getDisplayValues().filter(r=>r[0]===run):[];
 const nts=n&&n.getLastRow()>1?n.getRange(2,1,n.getLastRow()-1,11).getDisplayValues().filter(r=>r[0]===run):[];
 const agentByRow=new Map(),ntsByRow=new Map();
 // 사업자번호가 중복되더라도 원본행으로 구분해야 하므로 순서가 아닌 실행ID+사업자번호만으로 무조건 합치지 않음
 for(const r of agents){const k=String(r[3]).replace(/\D/g,'');if(!agentByRow.has(k))agentByRow.set(k,r);}
 for(const r of nts)ntsByRow.set(String(r[1]),r);
 const result=queue.map(r=>{
  const biz=String(r[3]).replace(/\D/g,''),agent=agentByRow.get(biz),tax=ntsByRow.get(String(r[1]));
  const ntsStatus=tax?(tax[4]==='조회완료'?tax[5]:'조회 '+tax[4]):'미조회';
  const dartStatus=r[8]==='완료'?(agent?agent[8]:'완료'):'처리 '+r[8];
  const signal=agent?agent[7]:'자료 없음';
  const completed=r[8]==='완료'||r[8]==='부분완료';
  const taxOk=tax&&tax[4]==='조회완료'&&String(tax[5]).startsWith('계속사업');
  const taxBlocked=tax&&tax[4]==='조회완료'&&(/휴업|폐업/.test(tax[5]));
  let score=0;
  if(taxOk)score+=40;
  if(completed)score+=15;
  if(agent&&agent[8]==='완료')score+=10;
  if(/매출채권 증가/.test(signal))score+=15;
  if(/제조|도매|건설|공사|용역|운송|유통|임대/.test(String(r[6])))score+=10;
  const priority=taxBlocked?'접촉보류':!taxOk?'검증대기':score>=70?'1차 검토':score>=55?'2차 검토':'추가 확인';
  const ready=completed&&tax&&tax[4]==='조회완료';
  return [run,p.getProperty('DAILY_DATE')||'',r[1],r[2],biz,ntsStatus,dartStatus,signal,'미연결·미검증',priority,'기초 우선순위(실제 미수채권 미확인) / 점수 '+score,ready?'검토용':'검증미완료'];
 });
 result.sort((x,y)=>{const rank={'1차 검토':0,'2차 검토':1,'추가 확인':2,'검증대기':3,'접촉보류':4};return (rank[x[9]]??5)-(rank[y[9]]??5)});
 const out=dailyTab_(DAILY.output,DAILY.headers);
 if(out.getLastRow()>1)out.getRange(2,1,out.getLastRow()-1,DAILY.headers.length).clearContent();
 if(out.getMaxRows()<result.length+1)out.insertRowsAfter(out.getMaxRows(),result.length+1-out.getMaxRows());
 out.getRange(2,1,result.length,DAILY.headers.length).setValues(result);
 const pending=queue.filter(r=>!['완료','부분완료','보류','실패'].includes(r[8])).length;
 const taxPending=nts.filter(r=>['대기','재시도'].includes(r[4])).length;
 const complete=pending===0&&taxPending===0&&nts.length===queue.length;
 p.setProperty('DAILY_READY',complete?'1':'0');
 if(complete){
  const archive=dailyTab_(DAILY.archive,DAILY.headers);
  const old=archive.getLastRow()>1?archive.getRange(2,1,archive.getLastRow()-1,1).getDisplayValues().some(x=>x[0]===run):false;
  if(!old){if(archive.getMaxRows()<archive.getLastRow()+result.length)archive.insertRowsAfter(archive.getMaxRows(),archive.getLastRow()+result.length-archive.getMaxRows());archive.getRange(archive.getLastRow()+1,1,result.length,DAILY.headers.length).setValues(result);}
 }
 dailyLog_('결과',run+' / '+result.length+'건 / DART 대기 '+pending+' / 국세청 대기 '+taxPending+' / '+(complete?'검토용 생성':'미완료'));
 return {run,companies:result.length,dartPending:pending,ntsPending:taxPending,allChecksFinished:complete,procurement:'미연결'};
}
function dailyStatus(){const p=PropertiesService.getScriptProperties();return {run:p.getProperty('DAILY_RUN'),inputDate:p.getProperty('DAILY_DATE'),count:p.getProperty('DAILY_SIZE'),ready:p.getProperty('DAILY_READY')==='1',note:'나라장터 통합은 미연결. 배포 전 확인 필요'};}
