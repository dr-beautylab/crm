# 뷰티샵 CRM v2 (클라우드플레어)

- `public/` 앱 화면 (기존 WebApp.html 기반)
- `functions/` 서버 (Pages Functions): `/api`, `/ingest`, `/cron/:job`, `/admin/import`
- `gas/Code.gs` 기존 Apps Script 로직 (개인정보 이관 데이터는 비움) → `node tools/gen.mjs gas/Code.gs` 로 `functions/_gas.js` 생성
- `gas/bridge.gs` 구글 쪽에 추가할 연결 스크립트 (메일 전달·시계·데이터 이관)
- 저장소: D1 (`DB` 바인딩), 비밀값: `CRM_KEY`
