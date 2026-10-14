# 단어 원고 · 검수 · MP3 통합 도구

## 내 컴퓨터에서 실행 (MP3 기능까지 쓰려면 이 방법 권장)
1. 이 폴더를 VS Code로 열기 (파일 → 폴더 열기)
2. 터미널 → 새 터미널에서 한 번만 실행:
   python -m pip install -r requirements.txt
3. 실행:
   python -m streamlit run app.py
4. 브라우저가 자동으로 열립니다.

## 탭 구성
- ① 엑셀 만들기: 기존 기능 그대로 (워드/PDF → 템플릿 엑셀)
- ② 엑셀 검수: 띄어쓰기, mp3 파일명 규칙, 중복, 번호 순서, Prompt 행, 고정값, 엑셀 파일명, 원본 대조
  - 자동 수정 가능한 항목은 수정본 다운로드 (수정 셀에 차수별 바탕색 + 붉은 글씨)
- ③ MP3 만들기: vocabulary_sound 열 파일명 그대로 생성 → 개수/파일명/용량/비트레이트/길이 검수 → ZIP

## 참고
- Streamlit Cloud 같은 서버에 배포하면 MP3 생성이 서버 IP 차단으로 실패할 수 있습니다.
- MP3가 전부 403 에러로 실패하면: python -m pip install --upgrade edge-tts
