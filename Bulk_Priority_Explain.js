/** 점수 산출 근거 상세표 V2. 기존 100개 점수표·원본은 보존. */
function bulkPriorityExplain100(){return bulkPriorityExplain_('20261025-112317');}
function bulkPriorityExplain_(run){
 const ss=SpreadsheetApp.openById(NIGHT.ssid),src=ss.getSheetByName('영업점수_'+run);
 if(!src)throw Error('기존 점수표 없음: '+run);
 const name='영업점수근거_'+run;
 if(ss.getSheetByName(name))throw Error('이미 생성됨: '+name+' (원본 보존)');
 const values=src.getDataRange().getValues(),header=values.shift();
 const pos=n=>header.indexOf(n);
 const idx={name:pos('기업명'),biz:pos('사업자등록번호'),total:pos('영업 우선검토 점수(100점 중)'),possible:pos('평가 가능 배점'),group:pos('검토구분'),basis:pos('재무기준'),r24:pos('2024 채권 관련 잔액(원)'),r25:pos('2025 채권 관련 잔액(원)'),note:pos('판정 근거·미확인 사항')};
 if(Object.values(idx).some(n=>n<0))throw Error('점수표 열 구조 불일치');
 const source=ss.getSheetByName('대량결과_'+run),base=source.getDataRange().getDisplayValues(),bh=base.shift();
 const bidx=n=>bh.indexOf(n),byBiz=new Map(base.map(r=>[String(r[bidx('사업자등록번호')]).replace(/\D/g,''),r]));
 const re=ss.getSheetByName('재무재검증_'+run),byRe=new Map();
 if(re&&re.getLastRow()>1)for(const r of re.getRange(2,1,re.getLastRow()-1,14).getValues())byRe.set(String(r[2]).replace(/\D/g,''),r);
 const data=[];
 for(const r of values){
  if(!r[idx.name])continue;
  const biz=String(r[idx.biz]).replace(/\D/g,''),orig=byBiz.get(biz),v=byRe.get(biz);
  const nts=orig?String(orig[bidx('국세청상태')]||''):'',dart=orig?String(orig[bidx('DART상태')]||''):'';
  const sN=nts.includes('계속사업')?5:0,sD=(dart==='완료'||(v&&v[5]==='완료 가능'))?5:0;
  let sSize=0,sGrowth=0,sRatio=0,delta='',ratio='',ruleSize='자료 미확인',ruleGrowth='자료 미확인',ruleRatio='자료 미확인';
  if(v&&v[5]==='완료 가능'){
   const rev24=Number(v[7]),rec24=Number(v[8]),rev25=Number(v[9]),rec25=Number(v[10]);
   if([rev24,rec24,rev25,rec25].every(Number.isFinite)&&rev24>0&&rev25>0&&rec24>=0&&rec25>=0){
    sSize=rec25>=1e10?10:rec25>=1e9?7:rec25>0?4:0;
    ruleSize=rec25>=1e10?'100억 원 이상':rec25>=1e9?'10억~100억 원 미만':rec25>0?'0~10억 원 미만':'0원';
    const d=rec24>0?(rec25-rec24)/rec24:null;
    delta=d===null?'전년 잔액 0':(d*100).toFixed(1)+'%';
    sGrowth=d!==null&&d>=.2?15:d!==null&&d>0?9:d!==null&&d===0?4:0;
    ruleGrowth=d===null?'전년 0원':d>=.2?'20% 이상 증가':d>0?'0~20% 미만 증가':d===0?'변동 없음':'감소';
    const q=rec25/rev25;ratio=(q*100).toFixed(1)+'%';
    sRatio=q>=.3?20:q>=.15?14:q>=.05?8:q>0?4:0;
    ruleRatio=q>=.3?'30% 이상':q>=.15?'15~30% 미만':q>=.05?'5~15% 미만':q>0?'0~5% 미만':'0%';
   }
  }
  const calc=sN+sD+sSize+sGrowth+sRatio;
  const total=Number(r[idx.total]);
  if(calc!==total)throw Error('점수 불일치 '+r[idx.name]+': '+calc+'/'+total+'; 결과표 생성 중단');
  data.push([r[idx.name],biz,total,Number(r[idx.possible]),r[idx.group],sN,sD,sSize,ruleSize,sGrowth,delta,ruleGrowth,sRatio,ratio,ruleRatio,0,0,0,'재무건전성·반복거래·접근성 미평가 45점',r[idx.basis],r[idx.note]]);
 }
 if(data.length!==values.filter(r=>r[idx.name]).length)throw Error('행 누락');
 const heads=['기업명','사업자등록번호','총점','평가 가능 배점','검토구분','국세청(5)','DART(5)','잔액 규모(10)','규모 점수 근거','잔액 증감(15)','실제 증감률','증감 점수 근거','매출 대비 비중(20)','실제 비중','비중 점수 근거','재무건전성(20)','반복거래(15)','영업접근성(10)','미평가 설명','재무기준','기존 판정 메모'];
 const sh=ss.insertSheet(name);
 sh.getRange(1,1,1,heads.length).setValues([heads]);
 for(let i=0;i<data.length;i+=150)sh.getRange(i+2,1,Math.min(150,data.length-i),heads.length).setValues(data.slice(i,i+150));
 sh.setFrozenRows(1);sh.setFrozenColumns(2);
 sh.getRange(1,1,1,heads.length).setBackground('#173451').setFontColor('#ffffff').setFontWeight('bold');
 sh.getRange(1,1,data.length+1,heads.length).createFilter();
 sh.setColumnWidth(1,185);sh.setColumnWidth(9,145);sh.setColumnWidth(12,155);sh.setColumnWidth(15,130);sh.setColumnWidth(19,265);sh.setColumnWidth(21,330);
 sh.getRange('C1').setNote('수임 확률이 아닌 영업 우선검토 점수. 최대 45점 미평가. 채권 잔액은 기타채권 포함 가능. 실제 연체 여부 미확인.');
 return {sheet:name,rows:data.length,matchedScores:true};
}
