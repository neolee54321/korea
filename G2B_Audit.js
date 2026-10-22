/**
 * 전수조회 사후 품질검사. 기존 전수매칭 및 원본 200개 기업은 수정하지 않습니다.
 * 실행 함수: g2bAuditFullScan
 */
function g2bAuditFullScan(){
 const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
 const progress=ss.getSheetByName('나라장터_전수진행').getDataRange().getDisplayValues();
 const last=progress[progress.length-1]||[];
 if(last[6]!=='complete')throw Error('전수조회 완료 기록이 없습니다');
 const matches=ss.getSheetByName('나라장터_전수매칭').getDataRange().getDisplayValues().slice(1);
 const day=String(last[1]);
 const key=PropertiesService.getScriptProperties().getProperty('G2B_API_KEY');
 if(!key)throw Error('G2B_API_KEY 없음');
 const ops=[['물품','getCntrctInfoListThng'],['공사','getCntrctInfoListCnstwk'],['용역','getCntrctInfoListServc'],['외자','getCntrctInfoListFrgcpt']];
 const results=[['검사일시','분야','API 현재 전체건수','최종페이지 반환건수','확인상태']];
 let total=0;
 ops.forEach(([category,op])=>{
   const base='https://apis.data.go.kr/1230000/ao/CntrctInfoService/'+op;
   const query=(page)=>{
     const q={ServiceKey:key,numOfRows:100,pageNo:page,type:'json',inqryDiv:1,inqryBgnDt:day+'0000',inqryEndDt:day+'2359'};
     const url=base+'?'+Object.keys(q).map(k=>k+'='+encodeURIComponent(q[k])).join('&');
     const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
     const json=JSON.parse(res.getContentText());
     const code=String(((json.response||{}).header||{}).resultCode);
     if(res.getResponseCode()!==200||!['0','00'].includes(code))throw Error(category+' API 응답 오류 '+code);
     return (json.response||{}).body||{};
   };
   const first=query(1);
   const count=Number(first.totalCount||0);
   const final=query(Math.max(1,Math.ceil(count/100)));
   const arr=Array.isArray(final.items)?final.items:((final.items||{}).item||[]);
   const n=Array.isArray(arr)?arr.length:arr?1:0;
   total+=count;
   results.push([new Date(),category,count,n,'마지막 페이지 재조회']);
 });
 const audit=ss.getSheetByName('나라장터_전수품질검사')||ss.insertSheet('나라장터_전수품질검사');
 audit.clearContents();
 audit.getRange(1,1,results.length,5).setValues(results);
 const seen={};let duplicates=0,missingName=0,futureDates=0;
 const today=Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd');
 matches.forEach(r=>{
   const k=[r[1],r[2],r[6]].join('|');
   if(seen[k])duplicates++;seen[k]=true;
   if(!r[5])missingName++;
   if(String(r[8]||'')>today)futureDates++;
 });
 const summary=[
   ['검사항목','결과','판정 참고'],
   ['전수조회 최종 상태',last[6],'진행 기록'],
   ['전수조회 처리건수',Number(last[4]||0),'실제 처리된 항목 수'],
   ['현재 API 전체건수 합계',total,'조회 시점 차이로 변동 가능'],
   ['처리건수 차이',total-Number(last[4]||0),'API 조회시점 변동 가능성 점검'],
   ['기존 200개 기업 일치 행',matches.length,'미수채권 확인 아님'],
   ['중복 매칭 행',duplicates,'분야·사업자번호·계약번호 기준'],
   ['계약명 누락 행',missingName,'원문 추가 검증 필요'],
   ['미래 계약일 행',futureDates,'조회 등록일과 계약일은 다를 수 있음']
 ];
 audit.getRange(8,1,summary.length,3).setValues(summary);
 return '품질검사 완료: 현재 전체 '+total+' / 처리 '+last[4]+' / 매칭 '+matches.length;
}
