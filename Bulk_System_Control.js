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
  if(ready&&typeof bulkFinalizeIfReady==='function')bulkFinalizeIfReady();
  const report={run,pending:pending.length,status:last[9],g2b:g2bStatus,ready,label};
  console.log(JSON.stringify(report));return report;
}
function bulkMonitorInstall() {
  const fn='bulkMonitorTick';
  const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===fn);
  if(!existing.length)ScriptApp.newTrigger(fn).timeBased().everyMinutes(30).create();
  return bulkMonitorTick();
}
