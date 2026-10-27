/**
 * 대량 영업 우선검토 점수표 (별도 시트 생성, 원본·일일 TOP10 불변).
 * 점수는 수임 확률이 아니라 공개 재무자료 기반의 상대적 검토 신호.
 * 실행: bulkPriorityPreview100() / 이후 bulkPriorityLatest().
 */
function bulkPriorityPreview100(){return bulkPriorityBuild_('20261025-112317');}
function bulkPriorityLatest(){
 const ss=SpreadsheetApp.openById(NIGHT.ssid),c=ss.getSheetByName('야간분석_실행관리');
 if(!c||c.getLastRow()<2)throw Error('실행관리 없음');
 const runs=c.getRange(2,1,c.getLastRow()-1,12).getDisplayValues().filter(r=>r[9]==='완료');
 for(let i=runs.length-1;i>=0;i--)if(ss.getSheetByName('대량결과_'+runs[i][0]))return bulkPriorityBuild_(runs[i][0]);
 throw Error('완료된 대량결과 시트 없음');
}
function bulkPriorityBuild_(run){
 const ss=SpreadsheetApp.openById(NIGHT.ssid),src=ss.getSheetByName('대량결과_'+run);
 if(!src)throw Error('원본 결과 없음: '+run);
 const data=src.getDataRange().getDisplayValues(),h=data.shift();
 const ix=n=>h.indexOf(n),ni=ix('기업명'),bi=ix('사업자등록번호');
 if(ni<0||bi<0)throw Error('원본 필수 열 없음');
 const base=data.filter(r=>r[ni]&&r[bi]);
 const re=ss.getSheetByName('재무재검증_'+run),lookup=new Map();
 if(re&&re.getLastRow()>1)for(const r of re.getRange(2,1,re.getLastRow()-1,14).getValues()){
  const k=String(r[2]).replace(/\D/g,'');if(k)lookup.set(k,r);
 }
 const rows=base.map((r,idx)=>{
  const name=r[ni],biz=String(r[bi]).replace(/\D/g,''),v=lookup.get(biz);
  const nt=String(r[ix('국세청상태')]||''),dart=String(r[ix('DART상태')]||'');
  let score=0,possible=0,notes=[],review='자료보완',basis='미확인',a=null,b=null,change='';
  // 자료 신뢰도 10점: NTS 상태 및 DART 확인. 실제 수임 가능성의 직접 증거 아님.
  possible+=10;
  if(nt.includes('계속사업')){score+=5;notes.push('계속사업 확인');}
  else notes.push('국세청 상태 재확인');
  if(dart==='완료'||(v&&v[5]==='완료 가능')){score+=5;notes.push('DART 재무조회 근거');}
  else notes.push('DART 근거 부족');
  // 채권잔액 규모·증감(25) 및 매출대비 비중(20): 비교 가능한 동일 재무기준 두 연도만 평가.
  if(v&&v[5]==='완료 가능'){
   const rev24=Number(v[7]),rec24=Number(v[8]),rev25=Number(v[9]),rec25=Number(v[10]);
   basis=String(v[6]||'미확인');a=rec24;b=rec25;
   if([rev24,rec24,rev25,rec25].every(Number.isFinite)&&rev24>0&&rev25>0&&rec24>=0&&rec25>=0){
    possible+=45;
    const delta=rec24>0?(rec25-rec24)/rec24:null;
    change=delta===null?'전년 0':(delta*100).toFixed(1)+'%';
    // 채권 잔액은 기타채권 합산 가능. 신호는 순위 참고용, 실제 미수 확인 아님.
    score+=rec25>=10000000000?10:rec25>=1000000000?7:rec25>0?4:0;
    score+=delta!==null&&delta>=.2?15:delta!==null&&delta>0?9:delta!==null&&delta===0?4:0;
    const ratio=rec25/rev25;
    score+=ratio>=.3?20:ratio>=.15?14:ratio>=.05?8:ratio>0?4:0;
    notes.push('동일 기준 2개 연도 비교','매출채권·기타채권 합산 가능');
   }else notes.push('재무 수치 비교 불가');
  }else notes.push(v?'재검증 '+v[5]:'재검증 미대상');
  // 미확인 영역 45점: 재무건전성 20, 반복거래 15, 접근성 10은 검증 없이 부여하지 않음.
  notes.push('재무건전성·반복거래·영업접근성 미평가(45점)');
  if(possible>=55)review=score>=40?'우선검토':'일반검토';
  // 55점은 확보 가능한 근거의 배점 합계, 100점으로 환산하거나 수임 확률로 해석 금지.
  return {idx,name,biz,score,possible,review,basis,a,b,change,notes:notes.join(' / '),source:String(r[ix('원본행')]||'')};
 });
 rows.sort((x,y)=>y.score-x.score||y.possible-x.possible||x.idx-y.idx);
 const title='영업점수_'+run;
 if(ss.getSheetByName(title))throw Error('기존 점수표 보존: '+title+' (중복 생성 중단)');
 const headers=['순위','기업명','사업자등록번호','영업 우선검토 점수(100점 중)','평가 가능 배점','검토구분','재무기준','2024 채권 관련 잔액(원)','2025 채권 관련 잔액(원)','채권 관련 잔액 증감률','판정 근거·미확인 사항','원본행'];
 const out=rows.map((r,i)=>[i+1,r.name,r.biz,r.score,r.possible,r.review,r.basis,r.a===null?'':r.a,r.b===null?'':r.b,r.change,r.notes,r.source]);
 if(out.length!==base.length)throw Error('입력·출력 기업 수 불일치');
 const sh=ss.insertSheet(title);
 sh.getRange(1,1,1,headers.length).setValues([headers]);
 if(out.length)for(let i=0;i<out.length;i+=200)sh.getRange(i+2,1,Math.min(200,out.length-i),headers.length).setValues(out.slice(i,i+200));
 sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setBackground('#173451').setFontColor('#ffffff').setFontWeight('bold');
 sh.getRange(1,1,out.length+1,headers.length).createFilter();
 sh.setColumnWidth(2,190);sh.setColumnWidth(11,460);
 sh.getRange(2,8,out.length,2).setNumberFormat('#,##0');
 sh.getRange(1,1).setNote('점수는 수임 확률(%)이 아닌 검토용 상대 지표. 평가 가능 배점도 함께 확인. 실제 연체·미수채권 미확인.');
 return {sheet:title,run,rows:out.length,completed:out.filter(x=>x[5]!=='자료보완').length,unscoredFinance:out.filter(x=>x[4]<55).length};
}

/** 기존 점수표는 유지하고, 세부 항목·계산 근거를 별도 시트에 기록합니다. */
function bulkPriorityExplain100(){
 const run='20261025-112317',ss=SpreadsheetApp.openById(NIGHT.ssid);
 const src=ss.getSheetByName('영업점수_'+run);
 if(!src)throw Error('선행 영업점수표가 없습니다');
 const rows=src.getDataRange().getValues(),head=rows.shift();
 const title='영업점수근거_'+run;
 if(ss.getSheetByName(title))throw Error('기존 근거표 보존: '+title);
 const detail=rows.filter(r=>r[1]).map(r=>{
  const [rank,name,biz,total,available,group,basis,prev,current,change,notes,original]=r;
  const f=financeByBiz.get(String(biz).replace(/\\D/g,''));
  const nts=String(notes).includes('계속사업 확인')?5:0;
  const dart=String(notes).includes('DART 재무조회 근거')?5:0;
  const a=typeof prev==='number'?prev:Number(prev),b=typeof current==='number'?current:Number(current);
  const financial=Number(available)===55&&Number.isFinite(a)&&Number.isFinite(b)&&a>=0&&b>=0;
  let scale='',growth='',ratio='',delta=null,ratioPercent='';
  if(financial){
   scale=b>=1e10?10:b>=1e9?7:b>0?4:0;
   delta=a>0?(b-a)/a:null;
   growth=delta!==null&&delta>=.2?15:delta!==null&&delta>0?9:delta!==null&&delta===0?4:0;
   ratio=Number(total)-nts-dart-scale-growth;
   if(![0,4,8,14,20].includes(ratio))throw Error('비중 점수 불일치: '+name+' '+ratio);
   const revenue=f?Number(f[9]):NaN;
   if(!Number.isFinite(revenue)||revenue<=0)throw Error('매출액 근거 누락: '+name);
   ratioPercent=(b/revenue*100).toFixed(1)+'%';
   const expected=b/revenue>=.3?20:b/revenue>=.15?14:b/revenue>=.05?8:b/revenue>0?4:0;
   if(expected!==ratio)throw Error('매출대비 비중 점수 재검산 불일치: '+name);
  }
  const sum=nts+dart+(financial?scale+growth+ratio:0);
  if(sum!==Number(total))throw Error('합계 검증 실패: '+name+' '+sum+'/'+total);
  const why=[
   '사업 '+nts+'/5: '+(nts?'계속사업 확인':'확인 근거 없음'),
   'DART '+dart+'/5: '+(dart?'조회 근거 있음':'조회 근거 부족'),
   financial?'잔액규모 '+scale+'/10: 2025 잔액 '+Math.round(b/1e8*10)/10+'억원':'잔액규모 미평가',
   financial?'증감 '+growth+'/15: '+(delta===null?'전년 0':(delta*100).toFixed(1)+'%'):'증감 미평가',
   financial?'매출대비 '+ratio+'/20: '+ratioPercent:'매출대비 미평가',
   '추가 45점 미평가: 건전성20·반복거래15·접근성10'
  ].join(' | ');
  return [rank,name,String(biz),total,available,group,nts,dart,scale,growth,ratio,0,why,basis,original,
   '재무항목은 기타채권 합산 가능. 점수는 수임확률이 아니며 미수·연체 여부 미확인'];
 });
 if(detail.length!==100)throw Error('100개 기업 불일치: '+detail.length);
 const h=['순위','기업명','사업자등록번호','총점','평가 가능 배점','검토구분','사업확인(5)','DART(5)','채권잔액규모(10)','전년증감(15)','매출대비비중(20)','미평가항목 점수(45)','점수 산출 근거','재무기준','원본행','주의사항'];
 const sh=ss.insertSheet(title);
 sh.getRange(1,1,1,h.length).setValues([h]);
 for(let i=0;i<detail.length;i+=100)sh.getRange(i+2,1,Math.min(100,detail.length-i),h.length).setValues(detail.slice(i,i+100));
 sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setBackground('#173451').setFontColor('#ffffff').setFontWeight('bold');
 sh.getRange(1,1,detail.length+1,h.length).createFilter();
 sh.setColumnWidth(2,180);sh.setColumnWidth(13,650);sh.setColumnWidth(16,300);
 return {sheet:title,rows:detail.length,verified:true};
}
