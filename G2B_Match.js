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
    let items=body.items&&body.items.item||[];
    if(!Array.isArray(items))items=[items];
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
