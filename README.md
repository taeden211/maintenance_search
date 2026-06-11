# 유지보수 사례 검색기

유지보수 내역서 엑셀 파일을 읽어서 장애내용/조치내용을 인덱싱하고, 다음 순서로 유사 사례를 찾는 GUI 프로그램입니다.

1. 키워드 검색: BM25
2. 벡터 검색: TF-IDF 기반 코사인 유사도
3. 조건 필터: 연도, 부서, 사용자, APC / PC filter / UTMP

## 준비 사항

- Windows
- Python 3
- 검색할 유지보수 내역서 Excel 파일(`.xlsx`)

Excel 데이터와 검색 인덱스는 저장소 및 실행 파일에 포함되지 않습니다.

## 소스에서 실행

저장소를 내려받은 폴더에서 다음 명령을 실행합니다.

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## 실행 파일 생성

다음 명령은 `.venv`가 없으면 자동으로 만들고 필요한 패키지를 설치한 뒤 실행 파일을 생성합니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

생성된 파일:

- `dist\유지보수사례검색기.exe`

다른 PC에는 이 실행 파일만 전달해도 프로그램을 실행할 수 있습니다. Python 설치는 필요하지 않지만, 검색할 Excel 파일은 별도로 준비해야 합니다.

## 처음 사용하는 방법

1. 프로그램을 실행합니다.
2. `찾기`를 눌러 유지보수 Excel 파일들이 들어 있는 최상위 폴더를 선택합니다.
3. `인덱스 구축`을 누릅니다.
4. 구축이 완료되면 검색어와 필터를 입력해 검색합니다.

인덱스는 선택한 데이터 폴더 아래의 `.maintenance_search_cache` 폴더에 저장됩니다. 같은 폴더를 다시 사용할 때는 `인덱스 불러오기`로 기존 인덱스를 사용할 수 있습니다. Excel 내용이 변경되면 `인덱스 구축`을 다시 실행하세요.
