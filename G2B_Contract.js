/**
 * 나라장터 계약정보 API 안전 진단용 모듈.
 * 실제 계약조회 오퍼레이션 경로는 공공데이터포털 활용신청 상세의
 * 요청주소를 확인한 뒤 설정해야 합니다. 임의 주소로 조회하지 않습니다.
 * 기존 TOP10, 국세청, DART 및 대량분석을 변경하지 않습니다.
 */
function g2bContractConfigCheck() {
  const p=PropertiesService.getScriptProperties();
  const key=String(p.getProperty('G2B_API_KEY')||'').trim();
  const url=String(p.getProperty('G2B_CONTRACT_ENDPOINT')||'').trim();
  const valid=/^https:\/\/apis\.data\.go\.kr\//.test(url);
  const status=!key?'인증키 미설정':!url?'계약정보 요청주소 미설정':!valid?'요청주소 확인 필요':'설정 준비 완료(조회 미실행)';
  const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
  const name='나라장터_계약정보진단';
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  sh.getRange(1,1,1,3).setValues([['점검시각','상태','다음 단계']]);
  sh.appendRow([new Date(),status,valid?'공식 문서 기준 파라미터 확정 후 시험조회':'공공데이터포털 계약정보서비스 상세의 요청주소 확인']);
  return status;
}
