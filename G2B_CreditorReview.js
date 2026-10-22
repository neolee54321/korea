/** 나라장터 일치기업 영업검토표: 원본 시트 불변, 공개 계약은 미수금 증빙이 아님. */
function g2bBuildCreditorReview(){
 const ss=SpreadsheetApp.openById('1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU');
 const src=ss.getSheetByName('대량분석_관리사배분').getDataRange().getDisplayValues();
 const match=ss.getSheetByName('나라장터_전수매칭').getDataRange().getDisplayValues().slice(1);
 const map={};src.slice(1).forEach(r=>{const no=String(r[3]||'').replace(/\D/g,'');if(no.length===10)map[no]=r;});
 const rows=[['사업자등록번호','기업명','업종','주소','대표전화','관리사','계약분야','계약번호','계약금액','계약명','계약일','검증상태','상담시 확인할 사항','주의']];
 match.forEach(r=>{
  const no=String(r[2]||'').replace(/\D/g,'');const a=map[no]||[];
  const issues=[];
  if(!r[5])issues.push('계약명 원문 확인');
  const today=Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd');
  if(String(r[8]||'')>today)issues.push('미래 계약일 원문 확인');
  const checks='외상매출·공사대금 미수 여부, 거래처 수, 연체기간, 증빙 보유 여부';
  rows.push([no,r[3]||a[2],a[6]||'',a[5]||'',a[4]||'',a[1]||r[4]||'',r[1],r[6],r[7],r[5],r[8],issues.length?issues.join(' / '):'계약정보 일치',checks,'계약정보는 미수채권 증빙이 아님']);
 });
 const name='나라장터_채권자영업검토';
 const sheet=ss.getSheetByName(name)||ss.insertSheet(name);
 sheet.clearContents();sheet.getRange(1,1,rows.length,rows[0].length).setValues(rows);
 sheet.setFrozenRows(1);
 return name+' 작성 완료: '+(rows.length-1)+'건';
}
