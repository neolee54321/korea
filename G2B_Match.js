/**
 * 나라장터 계약정보: 물품·공사·용역·외자 수집 및 사업자번호 대조.
 * 공개 계약정보만 조회하며 미수채권 보유 여부를 추정하지 않습니다.
 */
function g2bMatchContracts() {
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const source=ss.getSheetByName('대량분석_관리사배분');
  const rows=source.getDataRange().getDisplayValues().slice(1);
  const targets={};
  rows.forEach(r=>{
    const no=String(r[3]||'').replace(/\D/g,'');
    if(no.length===10)targets[no]={name:r[2],agent:r[1]};
  });
  const key=PropertiesService.getScriptProperties().getProperty('G2B_API_KEY');
  if(!key)throw Error('G2B_API_KEY 없음');
  const operations={
    '물품':'getCntrctInfoListThng',
    '공사':'getCntrctInfoListCnstwk',
    '용역':'getCntrctInfoListServc',
    '외자':'getCntrctInfoListFrgcpt'
  };
  const day=Utilities.formatDate(new Date(Date.now()-3*86400000),'Asia/Seoul','yyyyMMdd');
  const output=[['조회일','분야','사업자번호','업체명','담당관리사','계약명','계약번호','계약금액','계약일','확인수준']];
  const diagnostics=[];
  Object.keys(operations).forEach(category=>{
    const q={ServiceKey:key,numOfRows:100,pageNo:1,type:'json',inqryDiv:1,inqryBgnDt:day+'0000',inqryEndDt:day+'2359'};
    const base='https://apis.data.go.kr/1230000/ao/CntrctInfoService/'+operations[category];
    const url=base+'?'+Object.keys(q).map(k=>k+'='+encodeURIComponent(q[k])).join('&');
    const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
    let data={};try{data=JSON.parse(res.getContentText());}catch(e){}
    const h=(data.response||{}).header||{};
    const body=(data.response||{}).body||{};
    const code=String(h.resultCode==null?'미확인':h.resultCode);
    diagnostics.push([new Date(),category,res.getResponseCode(),code,body.totalCount||0]);
    if(res.getResponseCode()!==200||!['0','00'].includes(code))return;
    let items=Array.isArray(body.items)?body.items:(body.items&&body.items.item)||[];
    if(!Array.isArray(items))items=items?[items]:[];
    items.forEach(item=>{
      const list=String(item.corpList||'');
      const found=list.match(/\d{10}/g)||[];
      found.forEach(no=>{
        if(!targets[no])return;
        output.push([day,category,no,targets[no].name,targets[no].agent,
          String(item.cntrctNm||item.cntrctName||''),
          String(item.untyCntrctNo||''),
          String(item.totCntrctAmt||''),
          String(item.cntrctCnclsDate||''),
          '계약정보 일치·미수금 미확인']);
      });
    });
  });
  const dest=ss.getSheetByName('나라장터_계약매칭_시험')||ss.insertSheet('나라장터_계약매칭_시험');
  dest.clearContents();
  dest.getRange(1,1,output.length,10).setValues(output);
  const diag=ss.getSheetByName('나라장터_계약분야진단')||ss.insertSheet('나라장터_계약분야진단');
  diag.getRange(1,1,1,5).setValues([['검사시각','분야','HTTP','API코드','조회건수']]);
  diag.getRange(diag.getLastRow()+1,1,diagnostics.length,5).setValues(diagnostics);
  return '조회분야 4개 / 사업자번호 일치 '+(output.length-1)+'건 (1일, 분야별 최대 100건)';
}

/** 계약 API의 업체정보 필드 구조만 확인합니다. 실제 값과 인증키는 기록하지 않습니다. */
function g2bInspectContractFields(){
 const key=PropertiesService.getScriptProperties().getProperty('G2B_API_KEY');
 if(!key)throw Error('G2B_API_KEY 없음');
 const day=Utilities.formatDate(new Date(Date.now()-3*86400000),'Asia/Seoul','yyyyMMdd');
 const base='https://apis.data.go.kr/1230000/ao/CntrctInfoService/getCntrctInfoListThng';
 const q={ServiceKey:key,numOfRows:1,pageNo:1,type:'json',inqryDiv:1,inqryBgnDt:day+'0000',inqryEndDt:day+'2359'};
 const url=base+'?'+Object.keys(q).map(k=>k+'='+encodeURIComponent(q[k])).join('&');
 const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
 let j={};try{j=JSON.parse(res.getContentText())}catch(e){}
 const body=(j.response||{}).body||{};
 const items=body.items||{};
 const entry=Array.isArray(items)?items[0]:(Array.isArray(items.item)?items.item[0]:items.item);
 const keys=entry&&typeof entry==='object'?Object.keys(entry):[];
 const fields=keys.filter(k=>/corp|biz|brno|reg|cntrct|amt/i.test(k));
 const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
 const sh=ss.getSheetByName('나라장터_응답구조진단')||ss.insertSheet('나라장터_응답구조진단');
 sh.getRange(1,1,1,5).setValues([['검사시각','HTTP','응답구조','필드명(값 제외)','조회일']]);
 sh.appendRow([new Date(),res.getResponseCode(),Array.isArray(items)?'items 배열':Array.isArray(items.item)?'item 배열':typeof items.item,fields.join(', ').slice(0,2000),day]);
 return '필드명 '+fields.length+'개 확인';
}

/** 응답값을 공개하지 않고 JSON 계층 구조와 건수만 기록하는 보조 진단 */
function g2bInspectContractEnvelope(){
 const key=PropertiesService.getScriptProperties().getProperty('G2B_API_KEY');
 if(!key)throw Error('G2B_API_KEY 없음');
 const day=Utilities.formatDate(new Date(Date.now()-3*86400000),'Asia/Seoul','yyyyMMdd');
 const q={ServiceKey:key,numOfRows:1,pageNo:1,type:'json',inqryDiv:1,inqryBgnDt:day+'0000',inqryEndDt:day+'2359'};
 const base='https://apis.data.go.kr/1230000/ao/CntrctInfoService/getCntrctInfoListThng';
 const url=base+'?'+Object.keys(q).map(k=>k+'='+encodeURIComponent(q[k])).join('&');
 const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
 let data={};try{data=JSON.parse(res.getContentText())}catch(e){}
 const shape=x=>x&&typeof x==='object'?Object.keys(x).join(', '):typeof x;
 const response=data.response||{};
 const body=response.body||{};
 const items=body.items||{};
 const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
 const sh=ss.getSheetByName('나라장터_응답계층진단')||ss.insertSheet('나라장터_응답계층진단');
 sh.getRange(1,1,1,8).setValues([['검사시각','HTTP','최상위키','response키','body키','items구조','totalCount','resultCode']]);
 sh.appendRow([new Date(),res.getResponseCode(),shape(data),shape(response),shape(body),shape(items),String(body.totalCount||''),String((response.header||{}).resultCode||'')]);
 return '응답 계층 기록 완료';
}

/** 매칭 불일치 원인 진단: 개인정보/사업자번호 원문을 저장하지 않습니다. */
function g2bDiagnoseMatchingKeys(){
 const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
 const source=ss.getSheetByName('대량분석_관리사배분');
 const rows=source.getDataRange().getDisplayValues();
 const headers=rows[0]||[];
 const counts=headers.map((h,i)=>[i+1,h,rows.slice(1).filter(r=>String(r[i]||'').replace(/\D/g,'').length===10).length]);
 const key=PropertiesService.getScriptProperties().getProperty('G2B_API_KEY');
 if(!key)throw Error('G2B_API_KEY 없음');
 const day=Utilities.formatDate(new Date(Date.now()-3*86400000),'Asia/Seoul','yyyyMMdd');
 const q={ServiceKey:key,numOfRows:10,pageNo:1,type:'json',inqryDiv:1,inqryBgnDt:day+'0000',inqryEndDt:day+'2359'};
 const url='https://apis.data.go.kr/1230000/ao/CntrctInfoService/getCntrctInfoListThng?'+Object.keys(q).map(k=>k+'='+encodeURIComponent(q[k])).join('&');
 const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
 const data=JSON.parse(res.getContentText());
 const body=(data.response||{}).body||{};
 const items=Array.isArray(body.items)?body.items:[];
 const details=items.slice(0,10).map(x=>{
   const v=x.corpList;
   const raw=typeof v==='string'?v:JSON.stringify(v||'');
   return [typeof v,Array.isArray(v)?'array':v&&typeof v==='object'?'object':'text',raw.length,(raw.match(/\d{10}/g)||[]).length,(raw.match(/\d{3}-\d{2}-\d{5}/g)||[]).length];
 });
 const sh=ss.getSheetByName('나라장터_매칭키진단')||ss.insertSheet('나라장터_매칭키진단');
 sh.clearContents();
 sh.getRange(1,1,1,3).setValues([['원본 열번호','원본 열제목','10자리 번호형식 셀 수']]);
 if(counts.length)sh.getRange(2,1,counts.length,3).setValues(counts);
 const start=counts.length+4;
 sh.getRange(start,1,1,5).setValues([['corpList 자료형','구조','문자수','연속 10자리 수','하이픈 포함 번호 수']]);
 if(details.length)sh.getRange(start+1,1,details.length,5).setValues(details);
 return '원본 열 '+counts.length+'개, 계약 표본 '+details.length+'건 진단 완료';
}
