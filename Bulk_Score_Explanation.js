/**
 * 항목별 점수 산출 근거를 별도 시트에 기록.
 * 기존 영업점수/대량결과/일일 TOP10을 변경하지 않음.
 * 실행 함수: bulkScoreExplain100()
 */
function bulkScoreExplain100(){
 const run='20261025-112317',ss=SpreadsheetApp.openById(NIGHT.ssid);
 const source=ss.getSheetByName('영업점수_'+run),re=ss.getSheetByName('재무재검증_'+run);
 if(!source||!re)throw Error('점수표 또는 재검증표 없음');
 const name='영업점수근거_'+run;
 if(ss.getSheetByName(name))throw Error('기존 점수근거표 보존: '+name);
 const src=source.getDataRange().getDisplayValues(),hdr=src.shift();
 const col=s=>hdr.indexOf(s),bi=col('사업자등록번호');
 if(bi<0||src.length!==100)throw Error('점수표 100건 또는 사업자번호 확인 필요');
 const details=new Map();
 if(re.getLastRow()>1)for(const r of re.getRange(2,1,re.getLastRow()-1,14).getValues())details.set(String(r[2]).replace(/\D/g,''),r);
 const headers=['순위','기업명','사업자등록번호','국세청(5)','DART(5)','채권잔액 규모(10)','채권잔액 증감(15)','매출 대비 잔액(20)','미평가(45)','합계(100)','평가 가능 배점','채권잔액(2024,원)','채권잔액(2025,원)','증감률','매출 대비 잔액 비중','규모 산출 근거','증감 산출 근거','비중 산출 근거','자료 해석 주의'];
 const original=ss.getSheetByName('대량결과_'+run),originalRows=original.getDataRange().getDisplayValues(),originalHeader=originalRows.shift();
 const origMap=new Map(originalRows.map(r=>[String(r[2]).replace(/\D/g,''),r]));
 const out=src.map(r=>{
  const biz=String(r[bi]).replace(/\D/g,''),o=origMap.get(biz),v=details.get(biz);
  if(!o)throw Error('원본 누락: '+biz);
  const nts=String(o[3]||''),dart=String(o[4]||'');
  const ntsScore=nts.includes('계속사업')?5:0;
  const dartScore=(dart==='완료'||v&&v[5]==='완료 가능')?5:0;
  let size=0,growth=0,ratio=0,prior='',current='',change='',pct='',sizeWhy='재무 비교자료 미확인',growthWhy='재무 비교자료 미확인',ratioWhy='재무 비교자료 미확인';
  let possible=10;
  if(v&&v[5]==='완료 가능'){
   const rev24=Number(v[7]),rec24=Number(v[8]),rev25=Number(v[9]),rec25=Number(v[10]);
   if([rev24,rec24,rev25,rec25].every(Number.isFinite)&&rev24>0&&rev25>0&&rec24>=0&&rec25>=0){
    possible=55;prior=rec24;current=rec25;
    size=rec25>=1e10?10:rec25>=1e9?7:rec25>0?4:0;
    sizeWhy='2025 잔액 '+(rec25/1e8).toFixed(1)+'억원 → '+size+'/10점';
    const delta=rec24>0?(rec25-rec24)/rec24:null;
    change=delta===null?'전년 잔액 0':(delta*100).toFixed(1)+'%';
    growth=delta!==null&&delta>=.2?15:delta!==null&&delta>0?9:delta!==null&&delta===0?4:0;
    growthWhy=change+' → '+growth+'/15점';
    const q=rec25/rev25;pct=(q*100).toFixed(1)+'%';
    ratio=q>=.3?20:q>=.15?14:q>=.05?8:q>0?4:0;
    ratioWhy=pct+' → '+ratio+'/20점';
   }
  }
  const sum=ntsScore+dartScore+size+growth+ratio;
  const saved=Number(r[col('영업 우선검토 점수(100점 중)')]);
  if(sum!==saved||possible!==Number(r[col('평가 가능 배점')]))throw Error('기존 점수와 불일치: '+r[1]+' '+sum+'/'+saved);
  return [r[0],r[1],biz,ntsScore,dartScore,size,growth,ratio,'미평가',sum,possible,prior,current,change,pct,sizeWhy,growthWhy,ratioWhy,'수임 확률 아님 / 매출채권·기타채권 합산 가능 / 실제 미수·연체 미확인'];
 });
 const sh=ss.insertSheet(name);sh.getRange(1,1,1,headers.length).setValues([headers]);
 sh.getRange(2,1,out.length,headers.length).setValues(out);sh.setFrozenRows(1);
 sh.getRange(1,1,1,headers.length).setBackground('#173451').setFontColor('#ffffff').setFontWeight('bold');
 sh.getRange(1,1,out.length+1,headers.length).createFilter();
 sh.setColumnWidth(2,180);sh.setColumnWidth(16,240);sh.setColumnWidth(17,210);sh.setColumnWidth(18,210);
 sh.getRange(2,12,out.length,2).setNumberFormat('#,##0');
 sh.getRange(1,1).setNote('원본 점수 100건과 행별 총점/평가 가능 배점 일치 검증. 미평가 45점은 0점이 아닌 미확인.');
 return {sheet:name,rows:out.length,verified:out.length,source:'영업점수_'+run};
}
