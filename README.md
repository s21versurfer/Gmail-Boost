# 이메일 머지 — HTML 메일 작성기

스프레드시트(Excel/CSV) 데이터를 HTML 템플릿에 머지하여 Gmail로 발송하는 웹 앱입니다.

## 기능
- `{{변수명}}` 형식의 머지 태그
- 리치 텍스트 편집기 (굵게, 색상, 표, 구분선 등)
- 이미지 업로드 및 URL 삽입
- Excel (.xlsx) / CSV 파일 지원
- 행 범위 지정 (시작 행 ~ 끝 행)
- 변수 ↔ 열 자동 매핑
- 수신자별 이메일 미리보기
- Gmail 작성창 열기 / HTML 파일 저장

---

## 실행 방법

### A. 로컬에서 바로 열기
`index.html` 파일을 브라우저에서 열면 됩니다.
(Chrome / Edge / Firefox 권장)

### B. GitHub Pages로 배포 (무료, URL 공유 가능)

1. **GitHub 저장소 만들기**
   - https://github.com 에서 로그인
   - 우상단 `+` → `New repository`
   - Repository name: `mail-merge` (원하는 이름)
   - `Public` 선택 → `Create repository`

2. **파일 업로드**
   - 저장소 메인 페이지에서 `Add file` → `Upload files`
   - 이 폴더의 모든 파일을 업로드 (폴더 구조 유지):
     ```
     index.html
     css/style.css
     js/app.js
     README.md
     ```
   - `Commit changes` 클릭

3. **GitHub Pages 활성화**
   - 저장소 상단 `Settings` 탭 클릭
   - 왼쪽 메뉴 `Pages` 클릭
   - `Branch` → `main` 선택 → `/ (root)` → `Save`
   - 1~2분 후 다음 URL로 접속 가능:
     ```
     https://{내GitHub아이디}.github.io/mail-merge/
     ```

4. **수정 후 반영**
   - 파일 수정 후 저장소에 다시 업로드하면 자동 반영됩니다.

---

## 폴더 구조
```
mail-merge/
├── index.html       # 메인 페이지
├── css/
│   └── style.css    # 스타일
├── js/
│   └── app.js       # 기능 로직
└── README.md        # 이 파일
```

## 참고
- 서버 불필요 — 순수 HTML/CSS/JS
- 외부 CDN: Tabler Icons, SheetJS (xlsx)
- Gmail 전송 시 텍스트 버전으로 전달됨 (Gmail 보안 정책)
- HTML 서식 그대로 보내려면 `HTML 저장` 후 Gmail에 붙여넣기
