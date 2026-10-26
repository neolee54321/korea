/** 기존 점수표를 변경하지 않고 100개 기업별 배점 근거를 별도 시트로 생성. */
function bulkPriorityDetail100(){
 const run='20261025-112317',ss=SpreadsheetApp.openById(NIGHT.ssid);
 const source=ss.getSheetByName('영업점수_'+run),original=ss.getSheetByName('대량결과_'+run),check=ss.getSheetByName('재무재검증_'+run);
 if(!source||!original||!check)throw Error('원본 시트 누락');
 const title='영업점수상세_'+run;
 if(ss.getSheetByName(title))throw Error('기존 상세 점수표 보존');
 const a=source.getDataRange().getValues().slice(1).filter(r=>r[1]);
 const om=new Map(original.getDataRange().getDisplayValues().slice(1).map(r=>[String(r[2]).replace(/\D/g,''),r]));
 const cm=new Map(check.getDataRange().getValues().slice(1).map(r=>[String(r[2]).replace(/\D/g,''),r]));
 const rows=a.map(r=>{
  const key=String(r[2]).replace(/\D/g,''),o=om.get(key),v=cm.get(key);
  if(!o)throw Error('기업 누락 '+key);
  const nts=String(o[3]),dart=String(o[4]),n=nts.includes('계속사업')?5:0,d=dart==='완료'||(v&&v[5]==='완료 가능')?5:0;
  let size=0,growth=0,ratio=0,snote='자료 부족',gnote='자료 부족',rnote='자료 부족';
  if(v&&v[5]==='완료 가능'){
   const [sales0,debt0,sales1,debt1]=[v[7],v[8],v[9],v[10]].map(Number);
   if([sales0,debt0,sales1,debt1].every(Number.isFinite)&&sales0>0&&sales1>0&&debt0>=0&&debt1>=0){
    size=debt1>=1e10?10:debt1>=1e9?7:debt1>0?4:0;
    snote=(debt1/1e8).toFixed(1)+'억원';
    const delta=debt0>0?(debt1-debt0)/debt0:null;
    growth=delta!==null&&delta>=.2?15:delta!==null&&delta>0?9:delta!==null&&delta===0?4:0;
    gnote=delta===null?'전년 잔액 0원':(delta*100).toFixed(1)+'%';
    const fraction=debt1/sales1;
    ratio=fraction>=.3?20:fraction>=.15?14:fraction>=.05?8:fraction>0?4:0;
    rnote=(fraction*100).toFixed(1)+'%';
   }
  }
  if(n+d+size+growth+ratio!==Number(r[3]))throw Error('총점 불일치 '+r[1]);
  return [r[0],r[1],key,r[3],r[4],r[5],n,nts,d,dart,size,snote,growth,gnote,ratio,rnote,0,'재무건전성 20·반복거래 15·접근성 10 미평가',v?String(v[5]):'재검증 미대상',v?String(v[11]||''):''];
 });
 if(rows.length!==100)throw Error('100개 기업 불일치: '+rows.length);
 const headers=['순위','기업명','사업자번호','총점/100','평가가능 배점','검토구분','국세청/5','국세청 근거','DART/5','DART 근거','잔액규모/10','잔액규모 근거','잔액증감/15','증감률','매출대비/20','매출대비 비중','미평가점수/45','미평가 사유','재검증 상태','채권계정 구분'];
 const sh=ss.insertSheet(title);
 sh.getRange(1,1,1,headers.length).setValues([headers]);
 sh.getRange(2,1,rows.length,headers.length).setValues(rows);
 sh.setFrozenRows(1);sh.setFrozenColumns(3);
 sh.getRange(1,1,1,headers.length).setBackground('#173451').setFontColor('white').setFontWeight('bold');
 sh.getRange(1,1,rows.length+1,headers.length).createFilter();
 sh.setColumnWidth(2,190);[8,10,12,14,16,18,19,20].forEach(i=>sh.setColumnWidth(i,180));
 sh.getRange(1,4).setNote('수임 확률이 아닌 검토용 점수. 0점에는 미확인 항목도 포함.');
 return {sheet:title,rows:rows.length};
}
