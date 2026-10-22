/**
 * 나라장터 계약정보 API 안전 진단용 모듈.
 * 실제 계약조회 오퍼레이션 경로는 공공데이터포털 활용신청 상세의
 * 요청주소를 확인한 뒤 설정해야 합니다. 임의 주소로 조회하지 않습니다.
 * 기존 TOP10, 국세청, DART 및 대량분석을 변경하지 않습니다.
 */
function g2bContractConfigCheck() {
  const p=PropertiesService.getScriptProperties();
  const key=String(p.getProperty('G2B_API_KEY')||'').trim();
  const url=String(p.getProperty('G2B_CONTRACT_ENDPOINT')||'https://apis.data.go.kr/1230000/ao/CntrctInfoService').trim();
  const valid=url==='https://apis.data.go.kr/1230000/ao/CntrctInfoService';
  const status=!key?'인증키 미설정':!url?'계약정보 요청주소 미설정':!valid?'요청주소 확인 필요':'설정 준비 완료(조회 미실행)';
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const name='나라장터_계약정보진단';
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  sh.getRange(1,1,1,3).setValues([['점검시각','상태','다음 단계']]);
  sh.appendRow([new Date(),status,valid?'공식 문서 기준 파라미터 확정 후 시험조회':'공공데이터포털 계약정보서비스 상세의 요청주소 확인']);
  return status;
}

function g2bContractProbe() {
  const key=PropertiesService.getScriptProperties().getProperty('G2B_API_KEY');
  if(!key) throw Error('G2B_API_KEY missing');
  const day=Utilities.formatDate(new Date(Date.now()-259200000),'Asia/Seoul','yyyyMMdd');
  const query={ServiceKey:key,numOfRows:1,pageNo:1,type:'json',inqryDiv:1,inqryBgnDt:day+'0000',inqryEndDt:day+'2359'};
  const base='https://apis.data.go.kr/1230000/ao/CntrctInfoService/getCntrctInfoListThng';
  const url=base+'?'+Object.keys(query).map(k=>k+'='+encodeURIComponent(query[k])).join('&');
  const response=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
  let data={};try{data=JSON.parse(response.getContentText());}catch(e){}
  const header=(data.response||{}).header||{};
  const fault=(data.OpenAPI_ServiceResponse||{}).cmmMsgHeader||{};
  const code=String(header.resultCode!=null?header.resultCode:fault.returnReasonCode||'unknown');
  const status=response.getResponseCode()===200&&(code==='00'||code==='0')?'normal':'check';
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const sh=ss.getSheetByName('나라장터_계약조회진단')||ss.insertSheet('나라장터_계약조회진단');
  sh.appendRow([new Date(),status,response.getResponseCode(),code,day]);
  return status;
}
