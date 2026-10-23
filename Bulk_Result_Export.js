/**
 * 대량 기업분석 결과표 생성 v1.
 * TOP10과 무관. 기업 간 중복 여부는 사용자 검토 대상.
 * 원본 시트 변경 없음. 실행ID별 결과 스냅샷 시트를 생성합니다.
 * 함수: bulkCreateResultSheet()
 */
function bulkCreateResultSheet() {
  const ss = SpreadsheetApp.openById(NIGHT.ssid);
  const source = ss.getSheetByName('일일_분석이력');
  if (!source || source.getLastRow() < 2) throw Error('일일_분석이력 데이터 없음');
  const all = source.getRange(2,1,source.getLastRow()-1,12).getDisplayValues();
  const control = ss.getSheetByName('야간분석_실행관리');
  if (!control || control.getLastRow()<2) throw Error('실행관리 데이터 없음');
  const runs = control.getRange(2,1,control.getLastRow()-1,12).getDisplayValues();
  const run = runs[runs.length-1][0];
  const selected = all.filter(r=>r[0]===run);
  if (!selected.length) throw Error('해당 실행ID의 분석 결과 없음: '+run);
  const title = '대량결과_'+run;
  if (ss.getSheetByName(title)) {
    const msg='이미 결과표가 존재합니다: '+title+' (기존 결과 보호를 위해 덮어쓰지 않음)';
    console.log(msg);return {run,sheet:title,created:false,rows:selected.length};
  }
  const headers=['실행ID','기업명','사업자등록번호','국세청상태','DART상태',
    '재무징후','나라장터검증','영업검토순위','판정근거','배포상태',
    '원본행','자료해석 주의'];
  const output=selected.map(r=>[r[0],r[3],r[4],r[5],r[6],r[7],
    r[8],r[9],r[10],r[11],r[2],
    '재무징후·계약일치만으로 실제 미수채권 또는 연체가 확인되지는 않음']);
  const sh=ss.insertSheet(title);
  try {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    const size=250;
    for(let i=0;i<output.length;i+=size) {
      const part=output.slice(i,i+size);
      sh.getRange(i+2,1,part.length,headers.length).setValues(part);
    }
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setBackground('#e8eef4').setFontWeight('bold');
    sh.getRange(1,1,output.length+1,headers.length).createFilter();
    sh.autoResizeColumns(1,Math.min(headers.length,5));
    const result={run,sheet:title,created:true,rows:output.length};
    console.log(JSON.stringify(result));return result;
  } catch(e) {
    // 부분 기록도 검토할 수 있도록 시트를 남기고 오류를 명시합니다.
    try{sh.getRange(1,headers.length+1).setValue('생성 중 오류: '+String(e.message||e));}catch(ignore){}
    throw e;
  }
}

/**
 * 안전한 후속 처리 진입점.
 * 분석 실행관리 상태가 '완료'인 경우에만 결과표를 생성합니다.
 * 진행 중·오류·보류 상태에서는 어떤 원본 데이터도 변경하지 않습니다.
 * 함수: bulkFinalizeIfReady()
 */
function bulkFinalizeIfReady() {
  const ss=SpreadsheetApp.openById(NIGHT.ssid);
  const sh=ss.getSheetByName('야간분석_실행관리');
  if(!sh || sh.getLastRow()<2) throw Error('실행관리 시트 없음');
  const records=sh.getRange(2,1,sh.getLastRow()-1,12).getDisplayValues();
  const row=records[records.length-1];
  const run=row[0], status=String(row[9]||'').trim();
  if(status!=='완료') {
    const result={run,created:false,status,reason:'분석 미종료: 결과표 자동 생성 보류'};
    console.log(JSON.stringify(result));return result;
  }
  const g=dailyG2BStatus();
  if(!g||g==='미시작'||g.run!==run||g.status!=='complete')return {run,created:false,reason:'나라장터 검증 완료 대기'};
  const result=bulkCreateResultSheet();
  console.log(JSON.stringify(result));return result;
}
