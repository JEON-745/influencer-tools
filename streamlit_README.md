# 인플루언서 서칭 대시보드 (Streamlit)

## 로컬에서 먼저 테스트하기

1. 이 폴더를 통째로 GitHub 저장소로 올리기 전에, 로컬에서 먼저 돌려보는 걸 추천합니다.
2. 파이썬 패키지 설치:
   ```
   pip install -r requirements.txt
   ```
3. `.streamlit/secrets.toml.example` 파일을 복사해서 같은 폴더에 `secrets.toml`로 저장하고,
   실제 유튜브 API 키를 넣으세요.
   ```
   cp .streamlit/secrets.toml.example .streamlit/secrets.toml
   ```
4. 실행:
   ```
   streamlit run app.py
   ```
5. 브라우저에서 `http://localhost:8501` 접속

## GitHub에 올릴 때 주의

- `.gitignore`에 `secrets.toml`이 포함되어 있어서, `git add .`를 해도 실제 키는 저장소에 올라가지 않습니다.
- `secrets.toml.example`만 올라가는 게 정상입니다 (실제 키 없는 빈 템플릿).

## Streamlit Cloud 배포

1. [share.streamlit.io](https://share.streamlit.io) 에서 GitHub 계정으로 로그인
2. "New app" → 이 저장소 선택 → main 파일로 `app.py` 지정 → Deploy
3. 배포 후 앱 화면 우측 하단 "Manage app" → **Settings → Secrets** 메뉴에서
   `secrets.toml.example`과 같은 형식으로 실제 키를 붙여넣기
   ```
   YOUTUBE_API_KEY = "실제 키"
   ```

## 지금 버전에서 되는 것 / 안 되는 것

- ✅ 유튜브 실시간 검색, 필터링, 저장한 리스트, 검색 히스토리, 엑셀 내보내기(기존 서식 그대로)
- ⏳ 인스타그램 연동은 아직 없음 (서드파티 API 계약 필요)
- ⏳ Gmail 연동(메일 히스토리 확인)은 다음 단계에서 추가 예정
