/**
 * 나라장터 연결 진단. 기존 TOP10 및 야간분석 트리거를 변경하지 않습니다.
 * 최초에는 API 인증·응답 구조만 검사합니다. 기업 조회/계약 집계는 하지 않습니다.
 */
function g2bConnectionTest() {
  const key=String(PropertiesService.getScriptProperties().getProperty('G2B_API_KEY')||'').trim();
  if(!key) throw Error('스크립트 속성 G2B_API_KEY가 없습니다.');
  const url='https://apis.data.go.kr/1230000/ao/UsrInfoService02/getUnptRsttCorpInfo02'
    +'?serviceKey='+encodeURIComponent(key)+'&pageNo=1&numOfRows=1&type=json';
  const response=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:true});
  const http=response.getResponseCode();
  const body=response.getContentText();
  let parsed=null;
  try { parsed=JSON.parse(body); } catch(e) {}
  const header=parsed&&parsed.response&&parsed.response.header;
  const resultCode=header&&header.resultCode!=null?String(header.resultCode):'확인불가';
  const resultMsg=header&&header.resultMsg?String(header.resultMsg):'';
  const status=http===200&&header&&(resultCode==='00'||resultCode==='0')?'API 정상 응답':'추가 점검 필요';
  const fault=parsed&&parsed.OpenAPI_ServiceResponse&&parsed.OpenAPI_ServiceResponse.cmmMsgHeader;
  const errorCode=fault?String(fault.returnReasonCode||fault.errMsg||'').slice(0,60):'';
  const errorText=fault?String(fault.returnAuthMsg||fault.returnReasonCode||fault.errMsg||'').slice(0,100):'';
  const shape=parsed?'JSON 상위키: '+Object.keys(parsed).slice(0,8).join(', '):'JSON 아님';
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const name='나라장터_연결진단';
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  sh.getRange(1,1,1,5).setValues([['검사시각','HTTP','API 결과코드','검사상태','참고']]);
  sh.appendRow([new Date(),http,resultCode,status,(errorCode+' / '+errorText+' / '+resultMsg+' / '+shape).slice(0,120)]);
  return status+' / HTTP '+http+' / API '+resultCode+' (인증키와 원문 응답은 저장하지 않음)';
}
