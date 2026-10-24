/**
 * 대량 기업분석 운영 통합 점검판 v1.
 * 읽기 전용: 기존 입력/분석/나라장터 작업을 시작·중지·수정하지 않습니다.
 * TOP10, 기업 중복검사와 무관합니다.
 * 실행 함수: bulkSystemStatus
 */
function bulkSystemStatus() {
  const ss = SpreadsheetApp.openById(NIGHT.ssid);
  const p = PropertiesService.getScriptProperties();
  const control = ss.getSheetByName('야간분석_실행관리');
  if (!control || control.getLastRow() < 2) throw Error('야간분석_실행관리 자료 없음');
  const records = control.getRange(2,1,control.getLastRow()-1,12).getDisplayValues();
  const current = records[records.length-1];
  const run = current[0];
  const queue = ss.getSheetByName('야간분석_작업대기열');
  const counts = {};
  if (queue && queue.getLastRow() > 1) {
    const values = queue.getRange(2,1,queue.getLastRow()-1,16).getDisplayValues();
    values.forEach(r => {
      if (r[0] !== run) return;
      const stage = r[8] || '(상태 공란)';
      counts[stage] = (counts[stage] || 0) + 1;
    });
  }
  const g2bRaw = p.getProperty('DAILY_G2B_STATE');
  let g2b = null;
  if (g2bRaw) {
    try { g2b = JSON.parse(g2bRaw); }
    catch(e) { g2b = {status:'상태 파싱 오류'}; }
  }
  const sheet = ss.getSheetByName('일일_나라장터매칭');
  const g2bRows = sheet ? Math.max(0,sheet.getLastRow()-1) : 0;
  const report = {
    checkedAt: Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss'),
    run, total:current[4], complete:current[5], partial:current[6],
    held:current[7], failed:current[8], runStatus:current[9],
    lastUpdated:current[3], queueStages:counts,
    g2b: g2b ? {
      run:g2b.run,status:g2b.status,processed:g2b.processed,
      matches:g2b.matches,category:g2b.category,page:g2b.page,
      updated:g2b.updated,error:g2b.error||''
    } : '미시작',
    g2bSheetRows:g2bRows
  };
  console.log(JSON.stringify(report,null,2));
  return report;
}

/** 대기열 미종료 원인 확인. 읽기 전용이며 재시도/완료처리를 하지 않습니다. */
function bulkPendingCheck() {
  const ss=SpreadsheetApp.openById(NIGHT.ssid);
  const control=ss.getSheetByName('야간분석_실행관리');
  const queue=ss.getSheetByName('야간분석_작업대기열');
  if(!control||!queue) throw Error('운영 시트 누락');
  const runs=control.getRange(2,1,control.getLastRow()-1,12).getDisplayValues();
  const current=runs[runs.length-1],run=current[0];
  const values=queue.getRange(2,1,queue.getLastRow()-1,16).getDisplayValues();
  const rows=values.filter(r=>r[0]===run);
  const terminal=['완료','부분완료','보류','실패'];
  const pending=rows.filter(r=>!terminal.includes(r[8]));
  const report={
    run:run,expected:Number(current[4]),found:rows.length,
    runStatus:current[9],pendingCount:pending.length,
    pending:pending.slice(0,30).map(r=>({
      company:r[2],stage:r[7],status:r[8],attempts:r[9],
      nextAttempt:r[10],error:r[11]
    }))
  };
  console.log(JSON.stringify(report,null,2));
  return report;
}

/**
 * 통합 분석 감시. 30분마다 상태를 점검하고 별도 기록 시트에 남깁니다.
 * 기존 실행관리/대기열/원본/분석이력은 수정하지 않습니다.
 * 보류 및 진행중 기업을 자동으로 완료 처리하지 않습니다.
 */
function bulkMonitorTick() {
  const ss=SpreadsheetApp.openById(NIGHT.ssid);
  const ctrl=ss.getSheetByName('야간분석_실행관리');
  const q=ss.getSheetByName('야간분석_작업대기열');
  if(!ctrl||!q||ctrl.getLastRow()<2) throw Error('분석 운영 시트 없음');
  const last=ctrl.getRange(ctrl.getLastRow(),1,1,12).getDisplayValues()[0];
  const run=last[0],total=Number(last[4]||0);
  const values=q.getLastRow()>1?q.getRange(2,1,q.getLastRow()-1,16).getDisplayValues():[];
  const rows=values.filter(r=>r[0]===run);
  const done=new Set(['완료','부분완료','보류','실패']);
  const pending=rows.filter(r=>!done.has(r[8]));
  const g2b=PropertiesService.getScriptProperties().getProperty('DAILY_G2B_STATE');
  let g2bStatus='미시작';
  try {if(g2b)g2bStatus=JSON.parse(g2b).status||'상태없음';}catch(e){g2bStatus='파싱오류';}
  const monitor=ss.getSheetByName('대량분석_운영모니터')||ss.insertSheet('대량분석_운영모니터');
  if(monitor.getLastRow()===0)monitor.appendRow(['확인시각','실행ID','전체','대기열건수','완료','부분완료','보류','실패','미종료','실행상태','나라장터','판정']);
  const states={};
  rows.forEach(r=>{states[r[8]]=(states[r[8]]||0)+1;});
  const ready=rows.length===total&&pending.length===0&&last[9]==='완료';
  const label=ready?'결과표 생성 가능':pending.some(r=>r[8]==='진행중')?'진행중 상태 확인 필요 - 자동 재처리 안 함':pending.length?'미종료 '+pending.length+'건 - 상태 확인 필요':'실행 종료 판정 대기';
  const stamp=Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
  const record=[stamp,run,total,rows.length,states['완료']||0,states['부분완료']||0,states['보류']||0,states['실패']||0,pending.length,last[9],g2bStatus,label];
  const lastRow=monitor.getLastRow();
  const prev=lastRow>1?monitor.getRange(lastRow,2,1,11).getDisplayValues()[0]:[];
  const current=record.slice(1).map(String);
  if(JSON.stringify(prev)!==JSON.stringify(current))monitor.appendRow(record);
  if(ready){
    const archive=ss.getSheetByName('일일_분석이력');
    const archived=archive&&archive.getLastRow()>1?archive.getRange(2,1,archive.getLastRow()-1,1).getDisplayValues().filter(r=>r[0]===run).length:0;
    if(archived!==total && typeof dailyFinalize==='function'){
      try{dailyFinalize();}catch(e){console.error('자동 결과 취합 보류: '+String(e.message||e));}
    }
    if(typeof bulkFinalizeIfReady==='function'){
      try{bulkFinalizeIfReady();}catch(e){console.error('자동 결과표 생성 보류: '+String(e.message||e));}
    }
  }
  const report={run,pending:pending.length,status:last[9],g2b:g2bStatus,ready,label};
  try{bulkInputWaitingStatus();}catch(e){console.error(String(e.message||e));}
  console.log(JSON.stringify(report));return report;
}
function bulkMonitorInstall() {
  const fn='bulkMonitorTick';
  const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===fn);
  if(!existing.length)ScriptApp.newTrigger(fn).timeBased().everyMinutes(30).create();
  return bulkMonitorTick();
}

function bulkInputWaitingStatus(){
 const p=PropertiesService.getScriptProperties();
 const input=SpreadsheetApp.openById(NIGHT.ssid).getSheetByName(NIGHT.input);
 if(!input)throw Error('신규입력 시트 없음');
 if(input.getLastRow()<2 || !input.getRange(2,2,input.getLastRow()-1,1).getDisplayValues().some(r=>String(r[0]).trim())){
  const ss=SpreadsheetApp.openById(NIGHT.ssid);
  const s=ss.getSheetByName('대량분석_입력대기현황')||ss.insertSheet('대량분석_입력대기현황');
  s.getRange(1,1,2,5).setValues([['확인시각','입력건수','신규자료','기존분석중','상태'],[new Date(),0,'아니오','아니오','입력 대기']]);
  return {count:0,state:'입력 대기'};
 }
 const snap=dailySnapshot_();
 const fresh=p.getProperty('DAILY_INPUT_HASH')!==snap.hash;
 const busy=p.getProperty('NIGHT_ACTIVE')==='1'||p.getProperty('NIGHT_NTS_ACTIVE')==='1';
 const state=!fresh?'신규 입력 없음':busy?'신규 입력 대기':'신규 입력 준비';
 const ss=SpreadsheetApp.openById(NIGHT.ssid);
 const s=ss.getSheetByName('대량분석_입력대기현황')||ss.insertSheet('대량분석_입력대기현황');
 s.getRange(1,1,2,5).setValues([['확인시각','입력건수','신규자료','기존분석중','상태'],[new Date(),snap.rows.length,fresh?'예':'아니오',busy?'예':'아니오',state]]);
 return {count:snap.rows.length,state};
}

/** 자동분석을 시작하지 않는 사전 점검. */
function bulkDryRun(){
 const snap=dailySnapshot_();
 const p=PropertiesService.getScriptProperties();
 const busy=p.getProperty('NIGHT_ACTIVE')==='1'||p.getProperty('NIGHT_NTS_ACTIVE')==='1';
 const fresh=p.getProperty('DAILY_INPUT_HASH')!==snap.hash;
 const g=dailyG2BStatus();
 return {inputCount:snap.rows.length,newInput:fresh,priorRunBusy:busy,procurement:g==='미시작'?'미시작':g.status,canStart:fresh&&!busy,mode:'읽기전용 - 분석 시작 안 함'};
}

/** 기존 실행을 변경하지 않는 신규 실행 차단 원인 진단. */
function bulkNextRunGate(){
 const p=PropertiesService.getScriptProperties();
 const status=bulkPendingCheck();
 const active=p.getProperty('NIGHT_ACTIVE')==='1';
 const ss=SpreadsheetApp.openById(NIGHT.ssid);
 const archive=ss.getSheetByName('미종료보존_'+status.run);
 const saved=archive&&archive.getLastRow()>1&&p.getProperty('BULK_ARCHIVED_RUN')===status.run;
 const ntsActive=p.getProperty('NIGHT_NTS_ACTIVE')==='1';
 const g=dailyG2BStatus();
 const g2bActive=!!(g&&g!=='미시작'&&g.status==='running');
 const blocked=active||ntsActive||g2bActive||(status.pendingCount>0&&!saved);
 const result={priorRun:status.run,unfinished:status.pendingCount,archived:!!saved,active,ntsActive,g2bActive,canStartNext:!blocked,action:blocked?'실행 상태 또는 보존본 확인 필요':'신규 입력 시 실행 가능'};
 console.log(JSON.stringify(result));
 return result;
}

/** 기존 미종료 실행의 원본 상태를 별도 시트에 보존한다. 실행상태는 변경하지 않는다. */
function bulkArchivePending(){
 const info=bulkPendingCheck(),ss=SpreadsheetApp.openById(NIGHT.ssid);
 const q=ss.getSheetByName(NIGHT.queue);
 const all=q.getRange(2,1,q.getLastRow()-1,16).getDisplayValues();
 const rows=all.filter(r=>r[0]===info.run&&!['완료','부분완료','보류','실패'].includes(r[8]));
 const name='미종료보존_'+info.run;
 if(ss.getSheetByName(name))return {sheet:name,created:false,rows:rows.length};
 const out=ss.insertSheet(name);
 out.getRange(1,1,1,16).setValues([q.getRange(1,1,1,16).getDisplayValues()[0]]);
 if(rows.length)out.getRange(2,1,rows.length,16).setValues(rows);
 return {sheet:name,created:true,rows:rows.length};
}

/** 승인된 기존 실행을 분리 보관한 후에만 신규 실행 차단을 해제한다. */
function bulkReleaseArchivedRun(){
 return nightLock_(()=>{
  const p=PropertiesService.getScriptProperties();
  const id=p.getProperty('NIGHT_RUN');
  if(id!=='20261022-185519')throw Error('승인된 기존 실행ID와 다릅니다.');
  const ss=SpreadsheetApp.openById(NIGHT.ssid);
  const saved=ss.getSheetByName('미종료보존_'+id);
  if(!saved||saved.getLastRow()<2)throw Error('미종료 원본 보존 시트가 없습니다.');
  const q=ss.getSheetByName(NIGHT.queue);
  const rows=q.getRange(2,1,q.getLastRow()-1,16).getDisplayValues();
  const pending=rows.filter(r=>r[0]===id&&!['완료','부분완료','보류','실패'].includes(r[8]));
  const backup=saved.getRange(2,1,saved.getLastRow()-1,16).getDisplayValues();
  if(pending.length!==backup.length||pending.some(r=>!backup.some(b=>b[0]===r[0]&&b[1]===r[1]&&b[3]===r[3])))throw Error('미종료 원본과 보존본이 일치하지 않습니다.');
  if(p.getProperty('NIGHT_NTS_ACTIVE')==='1')throw Error('국세청 조회 실행 중: 전환 중단');
  const g=dailyG2BStatus();
  if(g&&g!=='미시작'&&g.status==='running')throw Error('나라장터 조회 실행 중: 전환 중단');
  p.setProperty('NIGHT_ACTIVE','0');
  nightClearTick_();
  p.setProperty('BULK_ARCHIVED_RUN',id);
  console.log('기존 미종료 원본 보존 확인 / 신규 실행 차단 해제: '+id);
  return {archivedRun:id,preserved:pending.length,nextRunReady:true};
 });
}
