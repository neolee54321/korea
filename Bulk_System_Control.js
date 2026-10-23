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
      const stage = r[15] || '(상태 공란)';
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
