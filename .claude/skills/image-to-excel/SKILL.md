---
name: image-to-excel
description: 이미지 폴더에서 OCR로 텍스트를 추출하여 Excel 파일로 저장합니다. 사용자가 "이미지에서 텍스트 추출", "OCR", "이미지를 엑셀로", "스캔 파일 변환" 등을 요청할 때 사용하세요.
argument-hint: [이미지 폴더 경로] [출력 파일명(선택)] [언어(선택)]
allowed-tools: Bash, Read, Glob
---

# 이미지 → Excel 텍스트 추출 스킬

사용자가 제공한 이미지 폴더에서 OCR로 텍스트를 추출하고 Excel 파일로 저장합니다.

## 인수 파싱

$ARGUMENTS 를 다음과 같이 해석하세요:
- **첫 번째 인수**: 이미지 폴더 경로 (필수)
- **두 번째 인수**: 출력 Excel 파일 경로 (선택, 없으면 이미지 폴더 안에 자동 생성)
- **`--lang` 또는 세 번째 인수**: OCR 언어 코드 (선택, 기본값: `kor+eng`)

인수가 없으면 사용자에게 이미지 폴더 경로를 물어보세요.

## 실행 단계

### 1단계: 환경 확인

다음 명령어로 필수 도구가 설치되어 있는지 확인하세요:

```bash
python --version
pip show Pillow pytesseract openpyxl
tesseract --version
```

설치되지 않은 항목이 있으면 사용자에게 아래 안내를 출력하세요:

```
[설치 필요]
1. Python 라이브러리:
   pip install Pillow pytesseract openpyxl

2. Tesseract OCR 엔진:
   - Windows: https://github.com/UB-Mannheim/tesseract/wiki
   - Mac:     brew install tesseract tesseract-lang
   - Ubuntu:  sudo apt install tesseract-ocr tesseract-ocr-kor
```

환경이 정상이면 2단계로 진행하세요.

### 2단계: 이미지 파일 확인

지정한 폴더에서 이미지 파일 목록을 확인하세요:

```bash
ls -1 <폴더경로>/*.{jpg,jpeg,png,bmp,tiff,tif,gif,webp} 2>/dev/null | head -20
```

- 이미지가 없으면 사용자에게 폴더 경로가 올바른지 확인하도록 안내하세요.
- 이미지 개수를 사용자에게 알려주세요.

### 3단계: OCR 실행 및 Excel 저장

프로젝트 루트의 `image_to_excel.py` 를 사용하여 실행하세요:

```bash
# 기본 실행 (출력 파일 자동 생성)
python image_to_excel.py <이미지폴더>

# 출력 파일 지정
python image_to_excel.py <이미지폴더> -o <출력파일.xlsx>

# 언어 지정 (한국어+영어가 기본값)
python image_to_excel.py <이미지폴더> --lang kor+eng

# 영어만
python image_to_excel.py <이미지폴더> --lang eng

# 한국어+영어+일본어
python image_to_excel.py <이미지폴더> --lang kor+eng+jpn
```

`image_to_excel.py` 가 없으면 사용자에게 파일 위치를 확인하거나 경로를 지정해 달라고 하세요.

### 4단계: 결과 보고

실행이 완료되면 다음 정보를 사용자에게 보고하세요:

1. **저장된 Excel 파일 경로**
2. **처리 통계**:
   - 전체 이미지 수
   - 텍스트 추출 성공 수
   - 실패/빈 결과 수
3. **오류가 발생한 파일**이 있으면 파일명과 오류 내용을 목록으로 표시
4. 결과 Excel에는 두 개의 시트가 있음을 안내:
   - `OCR 결과`: 이미지별 추출 텍스트 상세
   - `요약`: 전체 처리 통계

## 언어 코드 참고

| 언어 | 코드 |
|------|------|
| 한국어 | `kor` |
| 영어 | `eng` |
| 일본어 | `jpn` |
| 중국어(간체) | `chi_sim` |
| 한국어+영어 (기본) | `kor+eng` |

## 지원 이미지 형식

`.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.tif`, `.gif`, `.webp`

## 오류 처리

| 오류 상황 | 대응 방법 |
|-----------|-----------|
| 폴더를 찾을 수 없음 | 경로가 올바른지 확인하도록 안내 |
| Tesseract 미설치 | 설치 방법 안내 (1단계 참조) |
| 라이브러리 미설치 | `pip install -r requirements.txt` 안내 |
| 언어 팩 없음 | 해당 언어 팩 설치 방법 안내 |
| 권한 오류 | 폴더 읽기/쓰기 권한 확인 안내 |
