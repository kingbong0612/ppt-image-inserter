# 세신샵 PPT 자동 생성기

89개 업체의 이미지를 자동으로 PowerPoint 프레젠테이션에 삽입하는 Python 스크립트입니다.

## 🎯 주요 기능

- **자동 업체 스캔**: 지정된 디렉토리에서 모든 업체 폴더를 자동으로 찾습니다
- **엑셀 순서 적용**: 엑셀 파일의 업체명 순서대로 PPT를 생성합니다
- **이미지 자동 수집**: 각 업체의 `업체/업체_*.jpg` 파일들을 수집합니다
- **슬라이드 자동 생성**: 업체별로 표지, 가격표, 인테리어 슬라이드를 자동 생성합니다
- **이미지 최적화**: 원본 비율을 유지하면서 슬라이드에 맞게 자동 조정합니다

## 📋 요구사항

- Python 3.6 이상
- 필요한 라이브러리:
  ```bash
  pip install python-pptx pillow openpyxl
  ```

## 📁 디렉토리 구조

이 프로젝트는 [naver-map-photo-downloader](https://github.com/kingbong0612/naver-map-photo-downloader)의 다운로드된 이미지를 사용합니다.

**필요한 폴더 구조:**

```
GitHub/
├── naver-map-photo-downloader/
│   └── downloads/                   # 이미지 폴더 (자동 인식)
│       ├── 온아1호/
│       │   └── 업체/
│       │       ├── 업체_001.jpg
│       │       └── ...
│       ├── 온아2호/
│       │   └── 업체/
│       │       └── ...
│       └── ... (총 89개 업체 폴더)
│
└── ppt-image-inserter/              # 이 프로젝트
    ├── ppt_image_inserter.py        # 메인 스크립트
    ├── 세신샵.pptx                  # 템플릿 PPT 파일
    ├── 리스트_네이버지도링크추가.xlsx # 업체 순서 엑셀 파일
    └── 세신샵_완성본.pptx           # 생성된 PPT (자동 생성)
```

## 🚀 사용 방법

### 1. 라이브러리 설치

```bash
pip install python-pptx pillow openpyxl
```

또는 requirements.txt 사용:

```bash
pip install -r requirements.txt
```

### 2. 파일 준비

`ppt-image-inserter` 폴더에 다음 파일들을 배치하세요:
- `세신샵.pptx` - 템플릿 PPT 파일
- `리스트_네이버지도링크추가.xlsx` - 업체 순서 엑셀 파일

**이미지는 별도 설정 불필요!** `naver-map-photo-downloader/downloads` 폴더를 자동으로 찾습니다.

### 3. 실행

```bash
python ppt_image_inserter.py
```

메뉴에서 선택:
- `1` - 샘플 모드 (처음 10개 업체만 생성)
- `2` - 전체 작업 (모든 업체 생성)

### 4. 결과 확인

생성된 PPT 파일:
- 샘플: `세신샵_완성본_샘플.pptx`
- 전체: `세신샵_완성본.pptx`

## 📊 슬라이드 구조

업체 순서는 엑셀 파일의 매장명 순서를 따릅니다.

각 업체마다 다음과 같은 슬라이드가 생성됩니다:

1. **표지 슬라이드** - 업체명 표시
2. **가격표 슬라이드** - 첫 번째 이미지
3. **인테리어 슬라이드들** - 나머지 이미지들

## 🎨 이미지 처리

- **자동 비율 조정**: 이미지의 원본 비율을 유지하면서 슬라이드에 맞게 자동 조정
- **중앙 정렬**: 모든 이미지는 슬라이드 중앙에 배치
- **최대 크기**: 가로 11인치, 세로 5.5인치 이내로 자동 조정

## 📝 파일 설명

- **`ppt_image_inserter.py`** - 메인 스크립트 파일
- **`README_PPT사용법.md`** - 상세한 사용 설명서
- **`세신샵.pptx`** - 템플릿 PPT 파일 (별도 준비 필요)
- **`리스트_네이버지도링크추가.xlsx`** - 업체 순서 정보 엑셀 파일 (별도 준비 필요)

## 🔧 커스터마이징

### 경로를 직접 지정하고 싶다면

`ppt_image_inserter.py` 파일의 256번째 줄 부근에서 경로를 수정할 수 있습니다:

```python
# 이미지 디렉토리 (기본값: naver-map-photo-downloader의 downloads 폴더)
BASE_IMAGE_DIR = r"C:\Users\user\Documents\GitHub\naver-map-photo-downloader\downloads"

# 또는 다른 경로로 변경:
# BASE_IMAGE_DIR = r"C:\Users\user\Downloads\서울\중구"
```

### 슬라이드 제목 변경

`add_shop_to_ppt` 메서드에서 `slide_title` 변수를 수정하세요:

```python
if "가격표" in os.path.basename(image_path) or idx == 0:
    slide_title = "가격표"
else:
    slide_title = "인테리어"
```

### 이미지 크기 조정

`add_shop_to_ppt` 메서드에서 최대 크기를 조정하세요:

```python
max_width = Inches(11)
max_height = Inches(5.5)
```

### 텍스트 스타일 변경

```python
paragraph.font.size = Inches(0.5)
paragraph.font.bold = True
```

## 🔍 문제 해결

### 업체 디렉토리를 찾을 수 없습니다
- 경로가 정확한지 확인하세요
- 폴더 구조가 예상 구조와 일치하는지 확인하세요
- Windows 경로는 `r"C:\경로"` 형식을 사용하세요

### 이미지 추가 실패
- 이미지 파일이 손상되지 않았는지 확인하세요
- 파일 권한을 확인하세요
- 이미지 파일 형식이 지원되는지 확인하세요 (JPG, JPEG, PNG)

### 메모리 부족
- 이미지가 너무 많은 경우, 배치 처리로 나누어 실행하세요
- 이미지 파일 크기를 줄여보세요

## 📄 라이선스

이 프로젝트는 자유롭게 사용하실 수 있습니다.

## 🤝 기여

버그 리포트나 기능 제안은 이슈로 등록해주세요.

## 📧 문의

문제가 발생하면 스크립트 실행 시 출력되는 로그를 확인하세요.
