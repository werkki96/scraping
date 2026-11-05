# 핫딜 목록 수집기

퀘이사존과 아카라이브의 핫딜 게시판 목록을 수집해 엑셀(.xlsx), JSON, JSONL 또는 CSV로 저장합니다. Python 웹 스크래핑 프레임워크인 [Scrapling](https://github.com/D4Vinci/Scrapling) 0.5를 사용합니다.

## 빠른 시작

이 PC는 설치가 끝나 있습니다. 이 폴더(`scraping`)에서 명령 프롬프트(cmd)를 열고 실행합니다.

```bat
conda activate scrapping
python hotdeal_scraper.py
```

두 사이트를 각각 1페이지부터 최대 4페이지까지, 페이지 사이에 3분씩 쉬면서 수집합니다. 결과는 `deals.json`에 저장하고 `deals.xlsx`로 변환합니다. 이전 실행의 `deals.json`은 시작할 때 `before.json`으로 이름이 바뀝니다. 사이트마다 다음 중 하나가 나오면 그 사이트만 멈춥니다(2-1).

- `before.json`에 있는 글 (같은 글 번호 또는 같은 제목)
- 작성된 지 24시간이 지난 글

다음과 같이 동작합니다.

- **처음 실행:** `before.json`이 없으므로 24시간이 지난 글이 나오거나 4페이지가 될 때까지 수집합니다. 길면 약 10분 걸립니다.
- **다음 실행부터:** 보통 1~2페이지에서 멈추고 바로 끝납니다.
- **1페이지만 빠르게:** `--pages 1`을 붙입니다.

## 목차

1. [설치](#1-설치)
2. [실행](#2-실행)
3. [결과 파일](#3-결과-파일)
4. [실행 결과 확인과 문제 해결](#4-실행-결과-확인과-문제-해결)
5. [사이트 구조가 바뀌었을 때](#5-사이트-구조가-바뀌었을-때)
6. [수집 규칙](#6-수집-규칙)
7. [현재 상태와 알려진 한계](#7-현재-상태와-알려진-한계)

## 파일 구성

| 파일 | 내용 |
| --- | --- |
| `hotdeal_scraper.py` | 수집기 본체. 사이트별 셀렉터는 상단 `SITES`에 있음 |
| `export_excel.py` | 수집 결과(JSON)를 엑셀로 변환 |
| `requirements.txt` | pip 의존성 (`scrapling[fetchers]==0.5`, `openpyxl==3.1.5`) |
| `environment.yml` | conda 환경 정의 (환경 이름 `scrapping`, Python 3.12) |
| `.vscode/settings.json` | VS Code가 conda 환경 `scrapping`과 cmd 터미널을 쓰도록 하는 설정. PC마다 경로가 달라 저장소에는 넣지 않음(1-3) |
| `deals.json`, `deals.xlsx` | 이번 실행 결과 (실행 시 생성) |
| `before.json` | 직전 실행 결과. 시작할 때 `deals.json`을 이 이름으로 옮기며, 멈추는 조건 비교에 씀 (실행 시 생성) |
| `html_dump/` | `--dump-html`로 저장한 원본 HTML (실행 시 생성) |

## 1. 설치

이 PC에는 이미 설치되어 있습니다. 다른 PC에서 새로 설정할 때만 진행합니다. 필요한 것은 Windows와 Miniconda(또는 Anaconda)입니다.

### 1-1. conda 환경 만들기

```bat
conda env create -f environment.yml
```

conda 환경 `scrapping`(Python 3.12)을 만들고 `requirements.txt`의 패키지를 pip로 설치합니다. Scrapling은 conda 채널에 없어서 pip로 설치합니다.

이미 `scrapping` 환경이 있으면 패키지만 설치합니다.

```bat
conda activate scrapping
python -m pip install -r requirements.txt
```

설치 확인: `0.5`가 출력되면 됩니다.

```bat
conda run -n scrapping python -c "import scrapling; print(scrapling.__version__)"
```

### 1-2. cmd에서 `conda activate` 쓰기

이 PC의 cmd는 이미 설정되어 있습니다. 다른 PC에서 `conda activate`가 동작하지 않으면 아래 명령을 한 번 실행하고 cmd 창을 새로 엽니다. cmd가 시작될 때 conda를 불러오도록 사용자 레지스트리(`HKCU\Software\Microsoft\Command Processor`의 `AutoRun`)를 설정하는 명령입니다.

```bat
conda init cmd.exe
```

활성화 없이 실행하려면 명령 앞에 `conda run -n scrapping`을 붙입니다.

```bat
conda run -n scrapping python hotdeal_scraper.py -o deals.xlsx
```

### 1-3. VS Code

`.vscode/settings.json`에 다음 두 가지가 설정되어 있습니다.

- **Python 인터프리터:** conda 환경 `scrapping`을 씁니다. 오른쪽 아래 상태 표시줄에 `scrapping` 또는 `3.12.16`이 보이면 적용된 것입니다. 다른 인터프리터가 보이면 `Ctrl+Shift+P` → **Python: Select Interpreter** → `scrapping`을 선택합니다.
- **터미널:** 이 폴더에서 새로 여는 터미널은 cmd로 열립니다.

다른 PC에서는 `python.defaultInterpreterPath`를 그 PC의 conda 경로로 바꿉니다.

## 2. 실행

모든 명령은 `conda activate scrapping` 후 이 폴더에서 실행합니다. 프롬프트 앞에 `(scrapping)`이 보이면 활성화된 것입니다.

```bat
python hotdeal_scraper.py [옵션]
```

옵션 없이 실행하면 두 사이트를 최대 4페이지까지 3분 간격으로 수집합니다. 결과는 `deals.json`에 저장하고 `deals.xlsx`로 변환합니다.

| 옵션 | 설명 | 기본값 |
| --- | --- | --- |
| `--site quasarzone` / `--site arca` | 수집할 사이트. 두 번 쓰면 둘 다 | 두 사이트 모두 |
| `--pages N` | 사이트별 최대 목록 페이지 수 | 4 |
| `--interval 초` | 한 페이지를 처리한 뒤 다음 페이지를 요청하기 전 대기 시간 | 180 (3분) |
| `--max-age-hours 시간` | 작성된 지 이 시간이 지난 글이 나오면 그 사이트 수집을 멈춤. 0이면 제한 없음 | 24 |
| `--keyword 단어` | 제목에 단어가 들어간 글만 저장. 여러 번 쓰면 하나라도 맞는 글을 저장. 대소문자 무시 | 필터 없음 |
| `-o 파일` | 결과 파일. 확장자로 형식이 정해짐: `.json`, `.jsonl`, `.csv`. `.xlsx`를 주면 같은 이름의 `.json`과 `.xlsx`로 저장 | `deals.json` |
| `--before 파일` | 시작할 때 이전 결과를 옮겨 둘 파일. 이미 있으면 덮어씀 | `before` + 결과 파일 확장자 (`before.json`) |
| `--no-before` | 이전 결과를 옮기지도 비교하지도 않음. 결과 파일은 덮어씀 | 꺼짐 |
| `--excel 파일` | 엑셀 파일 이름 | 결과 파일과 같은 이름의 `.xlsx` |
| `--no-excel` | 엑셀로 변환하지 않음 | 꺼짐 |
| `--dump-html` | 각 사이트 1페이지 원본 HTML을 `html_dump\`에 저장. 결과가 이상할 때 원인 확인용 | 꺼짐 |
| `--browser` | 일반 HTTP 요청 대신 브라우저로 페이지를 열어 수집. 별도 설치 필요(6-3) | 꺼짐 |

예시:

```bat
REM 1페이지만 빠르게
python hotdeal_scraper.py --pages 1

REM 최대 2페이지, 페이지 사이 1분
python hotdeal_scraper.py --pages 2 --interval 60

REM 퀘이사존만, qz.json / qz.xlsx로 저장 (이전 qz.json은 before.json으로)
python hotdeal_scraper.py --site quasarzone -o qz.json

REM 엑셀 파일 이름만 바꾸기 / 엑셀 없이 JSON만
python hotdeal_scraper.py --excel 오늘핫딜.xlsx
python hotdeal_scraper.py --no-excel

REM 최근 12시간 글만
python hotdeal_scraper.py --max-age-hours 12

REM 제목에 SSD 또는 모니터가 들어간 글만
python hotdeal_scraper.py --keyword SSD --keyword 모니터

REM 결과가 이상할 때: 원본 HTML도 함께 저장
python hotdeal_scraper.py --dump-html
```

실행이 끝나면 Scrapling이 요청 수, 응답 코드, 차단 수 같은 통계를 길게 출력합니다. 결과 건수는 마지막 줄의 `N건 저장 → 파일 경로`에서 확인합니다.

- **한 사이트에 문제가 생기면:** 한쪽 사이트가 차단되거나 오류가 나도 다른 사이트 결과는 저장됩니다. 문제가 생긴 사이트는 다음 페이지로 넘어가지 않습니다.
- **이전 결과:** 시작할 때 `deals.json`을 `before.json`으로 옮깁니다. `before.json`은 직전 1회분만 남고 그 전 것은 덮어씁니다. 이전 결과와 합치지는 않습니다.
- **엑셀에서 연 파일:** 열어 둔 상태로는 덮어쓸 수 없습니다. 이때 JSON은 저장되고 엑셀 변환만 실패하며, 변환 명령을 알려 줍니다. 파일을 닫고 그 명령을 실행합니다.

### 2-1. 페이지 이동과 멈추는 조건

두 사이트는 동시에, 서로 따로 진행합니다. 사이트마다 다음 순서로 동작합니다.

1. 1페이지를 수집합니다.
2. 그 페이지에 다음 중 하나가 있으면 그 사이트는 여기서 멈춥니다. 다른 사이트는 계속 진행합니다.
   - **`before.json`에 있는 글:** 사이트가 같고 글 번호나 제목이 같은 글. 이 글은 다시 저장하지 않습니다.
     단, 제목만 같고 `before.json`의 글보다 나중에 작성된 글(매주 올라오는 쿠폰 글 등)은 새 글로 저장하고 멈추지 않습니다. 작성 시각을 알 수 없거나 같으면 글 번호가 큰 쪽을 나중 글로 봅니다.
   - **기간이 지난 글:** 작성된 지 `--max-age-hours`(기본 24시간)가 지난 글. 이 글은 저장하지 않습니다.
   - 어느 경우든 그 페이지의 나머지 새 글은 저장합니다.
3. 없으면 `--interval`(기본 3분)만큼 기다렸다가 다음 페이지를 수집하고 2로 돌아갑니다.
4. `--pages`(기본 4)페이지까지 수집하면 멈춥니다. 글이 0건인 페이지, 차단된 페이지, 요청이 실패한 페이지에서도 멈춥니다.

실제 페이지 간격은 대기 시간에 기본 요청 간격(3초 이상)이 더해져 약 3분 5초입니다.

같은 실행 안에서 본 글은 멈추는 조건이 아닙니다. 3분 사이에 새 글이 올라오면 1페이지에서 본 글이 2페이지로 밀려 내려오는데, 이 글은 한 번만 저장하고 계속 진행합니다.

**작성 시각 판정**

- **아카라이브:** 글의 작성 시각(UTC)을 그대로 씁니다.
- **퀘이사존 `N분 전`, `N시간 전`:** 수집 시각에서 그만큼 뺀 시각으로 봅니다.
- **퀘이사존 `11.04` 같은 날짜 표기:** 퀘이사존은 24시간 안의 글은 `N시간 전`으로, 그보다 오래된 글은 날짜로 보여 줍니다. 그래서 날짜 표기 글은 24시간이 지난 글로 봅니다. `--max-age-hours`를 24보다 크게 잡으면 그날 23:59에 쓴 글로 봅니다.
- **시각을 알 수 없는 글:** 기간 판정을 하지 않고 저장합니다.

**이전 결과 (`before.json`)**

- **비교 대상:** 직전 실행의 결과 파일입니다. `--keyword`로 결과에서 빠진 글은 들어 있지 않습니다.
- **처음부터 다시 수집:** `before.json`과 `deals.json`을 지우면 다음 실행은 처음 실행처럼 24시간 또는 `--pages`까지 수집합니다.
- **비교 없이 한 번만:** `--no-before`를 씁니다. 파일을 옮기지도 비교하지도 않고 결과 파일을 덮어씁니다.

**중간에 멈추기:** Ctrl+C를 한 번 누르면 기다리던 중이라도 1초 안에 다음 페이지 수집을 멈춥니다. 그때까지 수집한 결과는 정상 저장됩니다. 두 번 누르면 즉시 강제 종료하며, 이 경우 결과가 저장되지 않을 수 있습니다. 이때도 이전 결과는 시작할 때 이미 `before.json`으로 옮겨져 있습니다.

## 3. 결과 파일

### 3-1. 엑셀 (.xlsx)

수집이 끝나면 자동으로 만들어집니다(`--no-excel`로 끔). 이미 저장한 JSON·JSONL을 따로 변환할 수도 있습니다. 출력 파일명을 생략하면 같은 이름의 `.xlsx`로 저장합니다.

```bat
python export_excel.py deals.json
python export_excel.py deals.json -o 오늘핫딜.xlsx
```

| 시트 | 내용 |
| --- | --- |
| 핫딜 | 두 사이트 글을 작성 시각 기준 최신순으로 합친 목록. 제목을 누르면 게시글이 열립니다. 첫 줄 고정과 필터가 걸려 있습니다. `before.json`에 있는 글은 저장하지 않으므로 모든 글이 `신규` 열에 초록색으로 표시됩니다. 종료된 딜과 블라인드 글은 회색으로 표시합니다 |
| 요약 | 수집 시각, 전체·새 글·종료 건수, 사이트별·분류별·판매처별 건수 |

엑셀에서는 보기 편하도록 값을 정리합니다. 원본 JSON은 바뀌지 않습니다.

| 열 | 정리 방식 |
| --- | --- |
| 가격(원) | 원화 가격만 숫자로 바꿔 정렬할 수 있게 합니다. 달러·엔 등 외화는 비워 두고 `가격 표기` 열에 원래 값을 둡니다 |
| 배송 | `배송비 0`, `무배`를 `무료`로 통일합니다 |
| 작성 시각 | 한국 시각. 아카라이브는 UTC 시각을 변환하고, 퀘이사존은 `9시간 전` 같은 표기를 수집 시각 기준으로 환산합니다. 시간 단위 표기라 최대 1시간 오차가 있습니다 |

### 3-2. JSON·JSONL·CSV 필드

| 필드 | 내용 |
| --- | --- |
| `site` | 퀘이사존 / 아카라이브 |
| `post_id` | 게시글 번호 |
| `title` | 제목. 아카라이브 제목 끝의 댓글 수 표기(`[5]`)는 제거 |
| `url` | 게시글 주소 |
| `status` | 종료된 딜은 `종료`, 진행 중이면 빈 값 |
| `category`, `price`, `store`, `delivery` | 분류, 가격, 판매처, 배송 정보. 사이트에 표시된 그대로 |
| `comments`, `views`, `votes` | 댓글·조회·추천 수. `1.2k`, `2,394` 같은 표기를 정수로 변환. 표시가 없으면 빈 값 |
| `posted` | 작성 시각. 사이트 표기 그대로 (퀘이사존 `11분 전`, 아카라이브 UTC 시각) |
| `collected_at` | 수집 시각 |
| `is_new` | `before.json`에 없던 글. 그런 글만 저장하므로 항상 `true` |

CSV는 엑셀에서 한글이 깨지지 않도록 UTF-8 BOM을 붙여 저장합니다. 한 번의 실행 안에서 중복된 글은 한 번만 저장합니다.

## 4. 실행 결과 확인과 문제 해결

### 4-1. 로그로 원인 판단

| 로그·결과 | 의미 | 할 일 |
| --- | --- | --- |
| `N건 저장` (N > 0) | 정상 | 없음 |
| `이전 결과에 있는 글이 나와 … 수집을 마칩니다` | 직전 실행 이후 새 글을 모두 따라잡음. 정상 | 없음. 비교 없이 수집하려면 `--no-before` |
| `24시간이 지난 글이 나와 … 수집을 마칩니다` | 하루 치를 모두 수집함. 정상 | 없음. 기간을 바꾸려면 `--max-age-hours` |
| `새 글 0건` | 직전 실행 이후 새 글이 없음 | 없음 |
| `추출된 글이 0건입니다` | 페이지는 받았지만 글 목록을 찾지 못함. 사이트 개편 가능성 | 5장 |
| 일부 칸만 계속 비어 있음 | 해당 항목의 셀렉터가 맞지 않음 | 5장 |
| `봇 차단 응답 감지` | Cloudflare 등이 자동 접근을 막음 | 6-2 |
| `Request disallowed by robots.txt` | robots.txt가 해당 경로의 수집을 금지 | 그 사이트는 수집하지 않음 |
| `요청 실패` | 네트워크 오류, 시간 초과 등 | 인터넷 연결 확인 후 재실행 |

### 4-2. 실행 환경 문제

| 증상 | 해결 |
| --- | --- |
| `ModuleNotFoundError: No module named 'scrapling'` 또는 `openpyxl` | `scrapping` 환경이 아닌 Python으로 실행한 경우입니다. `conda activate scrapping` 후 실행합니다 |
| `python`과 파일명 사이 띄어쓰기 누락 (`pythonhotdeal_scraper.py`) | `python hotdeal_scraper.py`처럼 띄어 씁니다 |
| `conda activate`가 동작하지 않음 | PowerShell 창이면 cmd 창을 엽니다. cmd인데도 안 되면 1-2의 `conda init cmd.exe`를 실행합니다 |
| `PermissionError`로 저장 실패 | 같은 이름의 파일이 엑셀에서 열려 있습니다. 닫고 다시 실행합니다 |
| VS Code에서 import 오류 밑줄 | 1-3에 따라 인터프리터를 `scrapping`으로 선택합니다 |
| 터미널의 한글이 `���`로 깨짐 | 출력을 파일이나 다른 프로그램으로 넘길 때 생길 수 있습니다. `set PYTHONIOENCODING=utf-8` 후 다시 실행합니다. 저장 파일은 영향 없습니다 |

## 5. 사이트 구조가 바뀌었을 때

사이트가 개편되면 결과가 0건이 되거나 일부 칸이 비게 됩니다. 실제로 2026-11-04에 퀘이사존 목록이 새 구조(`v2-list-row`)로 바뀌어 있어서 셀렉터를 고쳤습니다.

사이트별 셀렉터는 `hotdeal_scraper.py` 상단의 `SITES`에 있습니다. 각 항목은 후보 목록이고, 앞에서부터 차례로 시도해 처음 값이 나온 것을 씁니다.

```python
"quasarzone": {
    "row":   ["div.v2-list div.v2-list-row", ...],       # 글 한 줄
    "skip_row_class": ["v2-partner-row"],                # 이 class가 붙은 줄은 건너뜀 (광고)
    "link":  ["a.subject-link", ...],                    # 글 링크
    "title": ["a.subject-link", ...],
    "price": ["span.v2-list-row__price", ...],
    "store": ["span.v2-list-row__ship img::attr(alt)"],  # ::attr(속성명)으로 속성 값 읽기
    ...
    "id_regex": r"/qb_saleinfo/views/(\d+)",             # 링크에서 글 번호 추출
},
```

이 밖에 `skip_if_text`(특정 요소의 글자가 일치하는 줄 건너뛰기)와 `done_marker`(이 요소가 있으면 상태를 `종료`로 표시) 항목이 있습니다. 아카라이브 설정에서 사용합니다.

수정 순서:

1. `python hotdeal_scraper.py --dump-html`로 원본 HTML을 저장합니다.
2. `html_dump\<사이트>_page1.html`을 브라우저로 열고 F12 개발자 도구로 글 한 줄을 감싸는 요소를 찾습니다.
3. `row`부터 맞춥니다. `row`가 틀리면 전체가 0건입니다.
4. `link`와 `id_regex`를 맞춥니다. 링크에서 글 번호를 찾지 못한 줄은 버려집니다.
5. 나머지 항목을 맞춥니다. 이 항목들은 틀려도 빈 값이 될 뿐 글은 저장됩니다.
6. 다시 실행해 확인합니다.

직접 고치기 어려우면 `html_dump` 폴더의 HTML 파일을 Claude에게 주고 셀렉터를 맞춰 달라고 하면 됩니다.

## 6. 수집 규칙

### 6-1. 요청 방식

| 설정 | 값 |
| --- | --- |
| robots.txt | 준수 (`Disallow`, `Crawl-delay` 반영) |
| 같은 사이트 동시 요청 | 1개 |
| 같은 사이트 요청 간격 | 최소 3초. 응답이 느리거나 오류가 나면 자동으로 늘어남 (AutoThrottle) |
| 페이지 이동 간격 | `--interval` (기본 3분) + 위의 요청 간격 |
| 전체 동시 요청 | 2개 (사이트당 1개씩) |
| 차단 응답 | 재시도하지 않고 건너뜀 |
| HTTP 모드 요청 형태 | Chrome 브라우저와 같은 TLS·헤더 (`impersonate="chrome"`) |
| 수집하지 않는 글 | 공지, 광고(퀘이사존 파트너 핫딜, 아카라이브 `광고` 배지 글) |

`--pages`를 크게 잡거나 `--interval`을 줄이면 요청이 늘어납니다. 요청 간격 설정(`download_delay`)과 페이지 간격은 줄이지 않는 것을 권장합니다. 두 사이트의 robots.txt와 이용약관은 사용 전에 직접 확인합니다. 스크립트는 robots.txt를 지키지만 이용약관의 자동 수집 금지 조항까지 판단하지는 않습니다.

- https://quasarzone.com/robots.txt
- https://arca.live/robots.txt

### 6-2. 차단될 때

`봇 차단 응답 감지`가 나오면 그 사이트가 자동 접근을 막고 있다는 뜻입니다. `--dump-html`을 함께 쓰면 차단 화면이 `html_dump\<사이트>_page1_blocked.html`로 저장됩니다.

- 스크립트는 차단 화면을 자동으로 풀거나 우회하지 않습니다. Scrapling에는 Cloudflare 챌린지 해제 기능이 있지만 일부러 넣지 않았습니다. 사이트가 명시적으로 막는 장치를 우회하면 이용약관 위반과 법적 분쟁 소지가 커지기 때문입니다.
- `--browser`는 JavaScript로 그려지는 내용을 읽기 위한 옵션이며, 차단 우회 수단이 아닙니다.
- 차단되는 사이트는 RSS나 공식 API처럼 허용된 방법이 있는지 먼저 확인하는 것을 권장합니다.

### 6-3. `--browser` 모드 준비

브라우저 모드는 Chromium을 별도로 내려받아야 합니다(수백 MB). 일반 모드로 목록이 비어 있거나, 퀘이사존 조회수처럼 JavaScript가 채우는 값이 필요할 때만 설치합니다.

```bat
scrapling install
python hotdeal_scraper.py --browser --dump-html
```

이 모드는 아직 실행해 보지 않았습니다.

## 7. 현재 상태와 알려진 한계

### 7-1. 확인 현황 (2026-11-04)

| 항목 | 상태 |
| --- | --- |
| 실제 수집 | 두 사이트 1페이지 수집 확인 (09:43 실행, 퀘이사존 30건·아카라이브 26건) |
| 셀렉터 | 두 사이트 모두 저장된 원본 HTML로 검증. 이후 아카라이브 광고 글 제외·종료 표시 판별을 추가했고, 이 수정 후에는 실제 사이트에서 다시 실행하지 않음 |
| 엑셀 변환 | 실제 수집 결과 56건으로 확인 |
| 로컬 테스트 | 샘플 HTML과 테스트 서버로 robots.txt 확인, 수집, 키워드 필터, 파일 저장, 차단 감지 확인 |
| 페이지 이동·멈춤 | 테스트 서버로 확인 (대기 1초로 단축): 기록 없음 → 4페이지까지, 같은 기록으로 재실행 → 1페이지에서 멈춤, 3페이지에 기록된 글 → 3페이지에서 멈춤, 밀려 내려온 글은 중복 저장 안 함, 대기 중 중단 요청 → 1초 안에 멈추고 결과 저장 |
| `before.json`·24시간 제한 (2026-11-05) | 테스트 서버로 확인: `deals.json` → `before.json` 이동·덮어쓰기, 같은 글 번호 또는 같은 제목에서 그 사이트만 멈춤, 같은 제목이라도 나중에 작성된 글은 저장하고 계속 진행, 24시간 지난 글(날짜 표기·UTC 시각)에서 멈춤, 엑셀 자동 변환, `-o`·`--excel` 이름 변경. 실제 사이트에서는 아직 실행하지 않음 |
| `--browser` 모드 | 브라우저를 설치하지 않았고 실행해 보지 않음 |
| 실제 사이트 2페이지 이후 | 실행해 보지 않음. 2페이지 이후 주소 형식(`?page=2`, `?p=2`)도 실제로 확인하지 않음 |
| robots.txt·이용약관 | 직접 확인하지 않음 |

### 7-2. 알려진 한계

| 사이트 | 항목 | 내용 |
| --- | --- | --- |
| 퀘이사존 | 댓글·조회·추천 수 | 페이지를 연 뒤 JavaScript(`qz-count-hydrate.js`)가 최신 값으로 다시 채우는 숫자입니다. 일반 모드로 받은 값은 그 이전 값이라 실제와 다를 수 있습니다. 예: 9시간 된 글의 조회수가 1~3으로 나옴 |
| 퀘이사존 | 작성 시각 | `11분 전`, `9시간 전` 같은 상대 표기만 있습니다. 엑셀에서는 추정 시각으로 환산합니다 |
| 퀘이사존 | 블라인드 글 | `블라인드 처리된 글입니다.`라는 제목으로 수집됩니다. 가격·판매처는 남아 있습니다. 엑셀에서는 회색으로 표시합니다 |
| 아카라이브 | 댓글 수 | 댓글이 없는 글은 표시가 없어서 빈 값이 됩니다 |
| 아카라이브 | 종료 표시 | 실제 페이지에서 종료된 일반 딜로는 확인하지 못했습니다. 종료 표시가 붙은 글이 광고 하나뿐이어서, 같은 구조의 샘플로만 확인했습니다 |

## 참고

- Scrapling: https://github.com/D4Vinci/Scrapling (BSD-3-Clause, 문서: https://scrapling.readthedocs.io)
- Scrapling README는 사용자가 각국 법, 사이트 이용약관, robots.txt를 지켜야 한다고 명시합니다.
