/**
 * 기업별 점수 차이 상세 근거: 기존 영업점수 결과를 읽고 새 시트만 생성.
 * 원본 결과/일일 TOP10/입력 데이터 수정 없음.
 * 실행 함수: bulkScoreBreakdown100
 */
function bulkScoreBreakdown100(){
 const run='20261025-112317',ss=SpreadsheetApp.openById(NIGHT.ssid);
 const source=ss.getSheetByName('영업점수_'+run);
 const finance=ss.getSheetByName('재무재검증_'+run);
 if(!source||!finance)throw Error('기존 점수표 또는 재무 재검증 시트 없음');
 const title='점수근거_'+run;
 if(ss.getSheetByName(title))throw Error('기존 상세점수표 보존: '+title);
 const values=source.getDataRange().getValues();
 const header=values.shift(),idx=s=>header.indexOf(s);
 const n=idx('기업명'),biz=idx('사업자등록번호'),total=idx('영업 우선검토 점수(100점 중)'),available=idx('평가 가능 배점');
 if([n,biz,total,available].some(x=>x<0))throw Error('기존 점수표 열 구성 불일치');
 const old=values.filter(r=>r[n]&&r[biz]);
 const lookup=new Map();
 if(finance.getLastRow()>1)for(const r of finance.getRange(2,1,finance.getLastRow()-1,14).getValues()){
  const k=String(r[2]).replace(/\D/g,'');if(k)lookup.set(k,r);
 }
 const output=old.map((r,i)=>{
  const k=String(r[biz]).replace(/\D/g,''),v=lookup.get(k);
  const sum=Number(r[total]),cap=Number(r[available]);
  let state='자료보완',basis='미확인',scale='',growth='',ratio='',scalePts=0,growthPts=0,ratioPts=0,details='';
  if(v&&String(v[5])==='완료 가능'){
   const [rev24,rec24,rev25,rec25]=[v[7],v[8],v[9],v[10]].map(Number);
   if([rev24,rec24,rev25,rec25].every(Number.isFinite)&&rev24>0&&rev25>0&&rec24>=0&&rec25>=0){
    basis=String(v[6]);state='2개년 재무 비교';
    scalePts=rec25>=1e10?10:rec25>=1e9?7:rec25>0?4:0;
    const change=rec24>0?(rec25-rec24)/rec24:null;
    growthPts=change!==null&&change>=.2?15:change!==null&&change>0?9:change!==null&&change===0?4:0;
    const fraction=rec25/rev25;
    ratioPts=fraction>=.3?20:fraction>=.15?14:fraction>=.05?8:fraction>0?4:0;
    scale=(rec25/1e8).toFixed(1)+'억원 → '+scalePts+'/10점';
    growth=(change===null?'전년 0원':(change*100).toFixed(1)+'%')+' → '+growthPts+'/15점';
    ratio=(fraction*100).toFixed(1)+'% → '+ratioPts+'/20점';
    details='2025 채권 관련 잔액 '+(rec25/1e8).toFixed(1)+'억원 / 2025 매출 '+(rev25/1e8).toFixed(1)+'억원 / 기타채권 합산 가능';
   }
  }
  const verified=cap>=55,baseline=sum-scalePts-growthPts-ratioPts;
  if(baseline<0||baseline>10||baseline+scalePts+growthPts+ratioPts!==sum)throw Error('점수 재계산 불일치: '+r[n]);
  if(verified&&state!=='2개년 재무 비교')throw Error('평가 가능 배점 불일치: '+r[n]);
  if(!verified&&state==='2개년 재무 비교')throw Error('재검증과 점수표 불일치: '+r[n]);
  return [i+1,r[n],k,sum,cap,baseline,scalePts,growthPts,ratioPts,0,0,0,state,basis,scale,growth,ratio,details,'미평가 45점: 재무건전성20·반복거래15·접근성10 / 수임확률 아님',r[idx('검토구분')]];
 });
 if(output.length!==100||new Set(output.map(x=>x[2])).size!==100)throw Error('100개 기업 수량 또는 사업자번호 확인 필요');
 const heads=['순위','기업명','사업자등록번호','총점','평가 가능 배점','사업·공시 확인(10)','채권잔액 규모(10)','전년 대비 증감(15)','매출 대비 잔액(20)','재무건전성(20)','반복거래(15)','접근성(10)','자료상태','재무기준','잔액규모 산출근거','증감률 산출근거','잔액비중 산출근거','재무수치·주의사항','미평가 사유','검토구분'];
 const sh=ss.insertSheet(title);
 sh.getRange(1,1,1,heads.length).setValues([heads]);
 sh.getRange(2,1,output.length,heads.length).setValues(output);
 sh.setFrozenRows(1);sh.setFrozenColumns(3);
 sh.getRange(1,1,1,heads.length).setBackground('#173451').setFontColor('#ffffff').setFontWeight('bold');
 sh.getRange(1,1,output.length+1,heads.length).createFilter();
 sh.setColumnWidth(2,180);sh.setColumnWidths(15,4,245);
 sh.getRange(1,1).setNote('항목별 배점은 검증 가능한 자료 기준. 미평가 항목은 0점으로 표기하지만 부정적 평가를 뜻하지 않음. 매출채권과 기타채권 합산 가능.');
 return {sheet:title,rows:output.length,checked:output.filter(x=>x[12]==='2개년 재무 비교').length,unverified:output.filter(x=>x[12]!=='2개년 재무 비교').length};
}
