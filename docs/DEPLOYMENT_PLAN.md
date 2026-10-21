# 야간분석 전용 Apps Script 연결 계획

## 현재 검증된 상태
- GitHub 저장소 읽기·쓰기: 확인
- Google Apps Script 프로젝트 코드 읽기·쓰기: 미연결
- 자동 배포: 미구현

## 연결 대상
- 독립 Apps Script 프로젝트 ID: `1jKSz8ArGKdpyRdyA_SROjs1aDbujdr_amm0dEYxb6-Bqu9XdQTfQmpnm`
- 스프레드시트 ID: `1Why-2qkGKASeMCIhmlPzk6FFmOLkIm_adi9yz_mmeGU`
- 기존 TOP10 프로젝트 및 트리거: 변경 금지

## 안전한 연결 순서
1. 현재 독립 Apps Script 프로젝트의 기존 코드 전체를 먼저 GitHub로 가져와 보존한다.
2. Google OAuth 인증은 사용자 계정에서 수행하고 인증정보는 GitHub Secrets로만 관리한다. 채팅, 공개 저장소에 올리지 않는다.
3. `clasp pull`로 원본 동기화를 검증한다.
4. 작은 시험용 함수 변경을 GitHub에서 수행한다.
5. `clasp push`는 사용자 승인 후에만 실행하고, 변경된 코드가 Apps Script에 반영됐는지 확인한다.
6. 그 후에만 야간 실행·재시도·이어하기·속도 제한·재무 및 주석 검증을 구현한다.

## 주의
- `clasp push`는 대상 Apps Script 프로젝트의 코드 파일을 덮어쓸 수 있다.
- 기존 코드 원본을 백업하기 전에는 push 금지.
- GitHub 저장소가 공개 상태이므로 회사 DB, 고객 연락처, API 키, 토큰은 커밋 금지.
- 중복 제거는 구현하지 않는다.
- 매출채권 관련 재무징후와 실제 연체채권·수임 가능성을 동일시하지 않는다.
