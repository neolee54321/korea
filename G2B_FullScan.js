/**
 * 나라장터 계약정보 전수 대조: 1일/4개 분야, 실행당 최대 8페이지.
 * 독립 시트에만 기록. 미수금 발생 여부는 확인하지 않습니다.
 * 시작: g2bStartFullScan() / 상태: g2bFullScanStatus()
 */
function g2bStartFullScan() {
  const props=PropertiesService.getScriptProperties();
  if(!props.getProperty('G2B_API_KEY'))throw Error('G2B_API_KEY 설정 필요');
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const src=ss.getSheetByName('대량분석_관리사배분');
  if(!src)throw Error('대량분석_관리사배분 시트 없음');
  const valid=src.getDataRange().getDisplayValues().slice(1).filter(r=>String(r[3]||'').replace(/\D/g,'').length===10);
  if(!valid.length)throw Error('사업자등록번호를 확인할 수 없음');
  if(props.getProperty('G2B_FULLSCAN_STATE'))return '기존 작업 존재: g2bFullScanStatus 확인';
  const day=Utilities.formatDate(new Date(Date.now()-3*86400000),'Asia/Seoul','yyyyMMdd');
  const state={day:day,category:0,page:1,processed:0,matches:0,status:'running',updated:new Date().toISOString()};
  props.setProperty('G2B_FULLSCAN_STATE',JSON.stringify(state));
  const out=ss.getSheetByName('나라장터_전수매칭')||ss.insertSheet('나라장터_전수매칭');
  out.clearContents();
  out.appendRow(['조회일','분야','사업자등록번호','기업명','관리사','계약명','통합계약번호','계약금액','계약체결일','검증수준']);
  const log=ss.getSheetByName('나라장터_전수진행')||ss.insertSheet('나라장터_전수진행');
  log.clearContents();
  log.appendRow(['실행시각','조회일','분야','마지막 페이지','처리 누계','일치 누계','상태']);
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='g2bFullScanBatch').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('g2bFullScanBatch').timeBased().everyMinutes(10).create();
  return '시작 준비 완료: '+day+' / 10분 간격 자동 실행';
}
function g2bFullScanBatch(){
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(1000))return '다른 실행 진행 중';
 try{
  const props=PropertiesService.getScriptProperties();
  const raw=props.getProperty('G2B_FULLSCAN_STATE');
  if(!raw)return '시작 상태 없음';
  const state=JSON.parse(raw);
  if(state.status!=='running')return state.status;
  const key=props.getProperty('G2B_API_KEY');
  if(!key)throw Error('인증키 없음');
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const source=ss.getSheetByName('대량분석_관리사배분');
  const targets={};
  source.getDataRange().getDisplayValues().slice(1).forEach(r=>{
    const no=String(r[3]||'').replace(/\D/g,'');
    if(no.length===10)targets[no]={name:r[2],agent:r[1]};
  });
  const ops=[['물품','getCntrctInfoListThng'],['공사','getCntrctInfoListCnstwk'],['용역','getCntrctInfoListServc'],['외자','getCntrctInfoListFrgcpt']];
  const out=ss.getSheetByName('나라장터_전수매칭');
  const log=ss.getSheetByName('나라장터_전수진행');
  let calls=0;
  while(calls<8&&state.category<ops.length){
   const [category,operation]=ops[state.category];
   const q={ServiceKey:key,numOfRows:100,pageNo:state.page,type:'json',inqryDiv:1,inqryBgnDt:state.day+'0000',inqryEndDt:state.day+'2359'};
   const url='https://apis.data.go.kr/1230000/ao/CntrctInfoService/'+operation+'?'+Object.keys(q).map(k=>k+'='+encodeURIComponent(q[k])).join('&');
   const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
   let data;try{data=JSON.parse(res.getContentText())}catch(e){throw Error(category+' '+state.page+'페이지 JSON 해석 실패')}
   const response=data.response||{};
   const code=String((response.header||{}).resultCode);
   if(res.getResponseCode()!==200||!['0','00'].includes(code))throw Error(category+' '+state.page+'페이지 HTTP '+res.getResponseCode()+' API '+code);
   const body=response.body||{};
   let items=Array.isArray(body.items)?body.items:(body.items&&body.items.item)||[];
   if(!Array.isArray(items))items=items?[items]:[];
   const batch=[];
   items.forEach(item=>{
     const numbers=[...new Set(String(item.corpList||'').match(/\d{10}/g)||[])];
     numbers.forEach(no=>{
       if(!targets[no])return;
       batch.push([state.day,category,no,targets[no].name,targets[no].agent,String(item.cntrctNm||''),String(item.untyCntrctNo||''),String(item.totCntrctAmt||''),String(item.cntrctCnclsDate||''),'계약업체 번호 일치 / 미수금 미확인']);
     });
   });
   // 기록 성공 후에만 체크포인트를 전진시킵니다.
   if(batch.length)out.getRange(out.getLastRow()+1,1,batch.length,10).setValues(batch);
   state.processed+=items.length;
   state.matches+=batch.length;
   calls++;
   const total=Number(body.totalCount||0);
   if(!items.length||state.page*100>=total){
     state.category++;state.page=1;
   }else state.page++;
   state.updated=new Date().toISOString();
   props.setProperty('G2B_FULLSCAN_STATE',JSON.stringify(state));
  }
  if(state.category>=ops.length){
   state.status='complete';
   props.setProperty('G2B_FULLSCAN_STATE',JSON.stringify(state));
   ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='g2bFullScanBatch').forEach(t=>ScriptApp.deleteTrigger(t));
  }
  log.appendRow([new Date(),state.day,state.category<ops.length?ops[state.category][0]:'전체 완료',state.page,state.processed,state.matches,state.status]);
  return '처리 '+state.processed+'건 / 일치 '+state.matches+'건 / '+state.status;
 }catch(e){
  const props=PropertiesService.getScriptProperties();
  const raw=props.getProperty('G2B_FULLSCAN_STATE');
  if(raw){
   const state=JSON.parse(raw);
   state.status='paused';state.error=String(e).slice(0,180);
   props.setProperty('G2B_FULLSCAN_STATE',JSON.stringify(state));
  }
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='g2bFullScanBatch').forEach(t=>ScriptApp.deleteTrigger(t));
  throw e;
 }finally{lock.releaseLock()}
}
function g2bFullScanStatus(){
 const raw=PropertiesService.getScriptProperties().getProperty('G2B_FULLSCAN_STATE');
 return raw?JSON.parse(raw):'시작 전';
}
