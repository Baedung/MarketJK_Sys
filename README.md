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

## 네이버 메일 실제 발송 설정 (선택)

"메일 발송" 탭의 **네이버로 실제 발송** 버튼을 쓰려면 Vercel 프로젝트에 아래 환경변수를 추가해야 합니다.

1. 네이버 메일 접속 → 환경설정 → POP3/IMAP 설정에서 **IMAP/SMTP 사용**을 켭니다.
2. 네이버 계정에 2단계 인증이 켜져 있다면, 네이버 아이디 관리에서 **애플리케이션 비밀번호(앱 비밀번호)**를 발급받습니다. (2단계 인증이 없다면 로그인 비밀번호를 그대로 사용할 수 있습니다.)
3. Vercel 프로젝트 → Settings → Environment Variables 에서 추가:
   - `NAVER_EMAIL` : 발송에 사용할 네이버 이메일 주소 (예: `mycompany@naver.com`)
   - `NAVER_APP_PASSWORD` : 위에서 발급한 비밀번호(또는 앱 비밀번호)
4. 저장 후 **재배포(Redeploy)** 해야 환경변수가 적용됩니다.

환경변수가 없으면 "네이버로 실제 발송" 버튼을 눌렀을 때 안내 메시지가 뜨고, 대신 "메일 초안" 버튼(기본 메일 앱 열기, 첨부는 직접 추가)은 설정 없이도 항상 사용할 수 있습니다.

## 폴더 구조

```
├─ api/send-mail.js     # 네이버 SMTP로 발주서를 첨부해 실제 발송하는 서버리스 함수
├─ src/App.jsx          # 전체 앱 (통합양식 / 매입처 관리 / 발주서 생성 / 메일 발송)
├─ src/main.jsx
├─ src/index.css
└─ index.html
```
