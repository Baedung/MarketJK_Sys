# 발주서 자동 생성 시스템

매입처별 발주서를 통합 양식 기준으로 자동 매칭해 생성하고, 네이버 메일로 발송하는 웹앱입니다.

## 로컬에서 실행

```bash
npm install
npm run dev
```

브라우저에서 `http://localhost:5173` 접속.

이 버전은 매입처/양식 데이터를 브라우저의 localStorage에 저장합니다. 즉, 같은 브라우저에서는 새로고침해도 유지되지만 다른 브라우저·기기와는 공유되지 않습니다.

## Vercel에 배포하기

### 1) GitHub에 올리고 Vercel에 연결 (권장)

```bash
git init
git add .
git commit -m "init"
```

이후 GitHub에 새 저장소를 만들고 push한 다음, [vercel.com](https://vercel.com) → "Add New Project" → 해당 저장소 선택 → 그대로 Deploy를 누르면 됩니다. (Framework Preset은 Vite로 자동 인식됩니다.)

### 2) 또는 Vercel CLI로 바로 배포

```bash
npm install -g vercel
vercel
```

질문에 답하면 배포되고, 이후 변경사항은 `vercel --prod`로 반영합니다.

## 네이버 메일 실제 발송 설정

앱 안의 **"설정" 탭**에서 네이버 이메일·앱 비밀번호를 입력·저장하면, "메일 발송" 탭의 **네이버로 실제 발송** 버튼이 그 계정으로 메일을 보냅니다. 이 설정이 실제로 저장되게 하려면 Vercel의 무료 저장소(KV)를 딱 한 번 연결해야 합니다.

1. Vercel 프로젝트 페이지 → **Storage** 탭 → **Create Database** → **KV** 선택 → 이름 정하고 생성.
2. 생성 화면에서 방금 만든 KV를 **이 프로젝트에 Connect**. (자동으로 `KV_REST_API_URL`, `KV_REST_API_TOKEN` 등 필요한 환경변수가 프로젝트에 추가됩니다 — 직접 입력할 필요 없음)
3. **Deployments** 탭에서 한 번 재배포(Redeploy).
4. 이후 앱의 "설정" 탭에서 네이버 이메일과 앱 비밀번호를 입력·저장하면 바로 적용됩니다. (다시 Vercel 대시보드에 갈 필요 없음)

네이버 쪽 준비:
- 네이버 메일 → 환경설정 → POP3/IMAP 설정에서 **IMAP/SMTP 사용**을 켭니다.
- 2단계 인증을 쓰는 계정이면, 네이버 아이디 관리에서 **애플리케이션 비밀번호**를 발급받아 그 값을 입력합니다. (2단계 인증이 없다면 로그인 비밀번호를 그대로 입력해도 됩니다.)

KV를 연결하지 않아도, 기존 방식대로 Vercel 프로젝트 Settings → Environment Variables에 `NAVER_EMAIL` / `NAVER_APP_PASSWORD`를 직접 추가하는 방법도 계속 동작합니다 (KV 설정이 없을 때 이 환경변수를 대신 사용합니다).

## 폴더 구조

```
├─ api/send-mail.js     # 네이버 SMTP로 발주서를 첨부해 실제 발송하는 서버리스 함수
├─ api/settings.js       # '설정' 탭의 네이버 계정 정보를 Vercel KV에 저장/조회
├─ src/App.jsx          # 전체 앱 (통합양식 / 매입처 관리 / 발주서 생성 / 메일 발송 / 설정)
├─ src/main.jsx
├─ src/index.css
└─ index.html
```
