# 관내출장비 집계 웹앱

GitHub에 올린 뒤 Streamlit Community Cloud에서 실행할 수 있는 PDF 출장비 계산기입니다.

## 로컬 실행

1. 인사랑에서 관내출장 결재내역을 PDF로 내려받습니다.
2. 필요한 패키지를 설치합니다.

```powershell
python -m pip install -r requirements.txt
```

3. 웹 서버를 실행합니다.

```powershell
streamlit run main.py
```

브라우저에서 PDF를 업로드하고 `PDF이름=실제이름`을 입력하면 사람별 출장 횟수와 총액을 확인할 수 있습니다. 예를 들어 `코리요=장병순`으로 입력하면 PDF의 코리요 출장 건이 표준 지급명세서의 장병순 행으로 합산됩니다. `사무분장표 순서`에는 실제 이름을 표시 순서대로 한 줄에 하나씩 입력합니다. 입력한 이름은 그 순서로 출력되고, 목록에 없는 사람은 PDF 순서로 뒤에 표시됩니다. `엑셀 다운로드` 버튼은 표준 서식의 완성 파일을 내려받습니다.

스캔 이미지 PDF는 `pymupdf`, `pytesseract`, `pillow`와 Tesseract OCR 한국어 언어팩이 필요합니다. OCR이 읽은 출장 행에서 `4시간 미만=10,000원`, `4시간 이상=20,000원`, 차량 사용 시 `10,000원 차감` 규칙으로 계산합니다. 이름이나 시간 인식이 틀린 경우에는 원본 PDF의 해상도를 높여 다시 내려받아 주세요.

## GitHub 배포

1. 이 폴더의 파일을 GitHub 저장소에 업로드합니다.
2. [Streamlit Community Cloud](https://share.streamlit.io/)에서 `Deploy an app`을 선택합니다.
3. GitHub 저장소와 브랜치를 선택하고 Main file path에 `main.py`를 입력합니다.
4. Deploy를 누르면 공개용 웹 주소가 생성됩니다.

개인정보가 포함된 결재 PDF와 실제 지급명세서 원본은 GitHub에 올리지 마세요. 이 프로젝트의 `.gitignore`가 원본 자료와 가상환경을 제외합니다.