# LumbarDISC 쉬모를 결절 라벨링 도구

RSNA 2024 Lumbar Spine 데이터의 Sagittal T2 영상을 보면서 쉬모를 결절(Schmorl node)을
엑셀 양식(`LumbarDISC_Schmorl_labeling_v0.1.xlsx`)에 기록하는 도구입니다.
판독은 MicroDicom으로 하고, 이 프로그램은 기록용입니다.

## 폴더 배치

모든 데이터와 결과물은 **`C:\Users\kay85\OneDrive\바탕 화면\연구멘토링`** 안에서만 읽고 씁니다.
기준 경로는 `config.py`의 `PROJECT_DIR`이고, `tool` 폴더를 어디서 실행하든 이 경로가 기준입니다.
관리하기 쉽도록 `tool` 폴더도 이 안에 두는 것을 권장합니다.

```
연구멘토링\
├─ data\                      ← RSNA 데이터 (train_images, CSV 2개 — CSV는 data 바로 아래 하위 폴더에 있어도 찾음)
├─ labeling\
│   ├─ LumbarDISC_Schmorl_labeling_v0.1.xlsx   ← 결과 엑셀
│   ├─ captures\              ← 애매(Uncertain) 병변 캡처 이미지
│   ├─ backup\                ← 엑셀 수정 전 자동 백업 (자동 생성)
│   ├─ setup_report.txt       ← ⓪ 점검 보고서
│   └─ sample_selection_log.txt ← ① 추출 기록 (시드, 모집단, 목록)
└─ tool\                      ← 이 폴더 (코드)
```

폴더 이름이나 위치가 바뀌면 `config.py`의 `PROJECT_DIR` 한 줄만 고치면 됩니다.
실행 파일(.bat)은 `__pycache__` 같은 부산물을 만들지 않도록 설정돼 있습니다.

## 처음 한 번: 준비

1. **Python 설치**: https://www.python.org/downloads/ 에서 3.10 이상을 설치합니다.
   설치 첫 화면에서 **"Add python.exe to PATH"를 꼭 체크**하세요.
2. **패키지 설치**: `tool\install.bat`을 더블클릭합니다. 끝에 `Successfully installed ...`가
   나오거나 이미 설치돼 있다는 메시지가 나오면 정상입니다.

## 단계별 실행

| 단계 | 실행 파일 | 하는 일 | 파일 수정 |
|---|---|---|---|
| ⓪ 점검 | `step0_check.bat` | 폴더·CSV·DICOM·엑셀 양식·MicroDicom 위치를 확인하고 보고서를 만듭니다 | 없음 (보고서만 저장) |
| ① 추출 | `step1_make_sample.bat` | 무작위 100명을 뽑아 `sample_list` 시트에 기록합니다 | 엑셀 (수정 전 자동 백업) |
| 정리 | `fix_remove_rater.bat` | 엑셀 양식에서 rater 열을 지웁니다 (한 번만) | 엑셀 (수정 전 자동 백업) |
| 문제 진단 | `diagnose_excel.bat` | "엑셀 읽기 실패"가 날 때 원인과 정상 백업을 찾습니다 | 없음 (보고서만 저장) |
| ②~⑥ 라벨링 | `run_labeler.bat` | 라벨링 프로그램 (단계마다 기능이 추가됨) | 엑셀·캡처 (④부터) |

엑셀을 수정하는 단계는 **실행 전에 엑셀 파일을 닫아 두세요.** 열려 있으면 저장이 실패하고,
원본은 그대로 남습니다.

### ⓪ 점검 결과 보는 법
- `[OK]`: 정상, `[없음]`: 경로나 파일을 확인해야 합니다.
- 보고서 파일은 `labeling\setup_report.txt`에 저장됩니다.
- **6. 엑셀 양식** 항목은 각 열의 머리글 색을 `회색(프로그램)`, `초록(수식)`, `노랑(사람)`으로
  구분해서 보여 줍니다. 프로그램은 **회색 열에만** 값을 씁니다.

### ① 추출 방식 (논문 방법 기술용)
- 모집단: `train_images`에 Sagittal T2/STIR 시리즈가 실제로 다운로드된 study
- study_id를 오름차순으로 정렬한 뒤 `random.Random(2024).sample(목록, 100)`으로 추출합니다.
  시드(2024)와 인원(100)은 `config.py`에서 바꿀 수 있습니다.
- 같은 데이터와 같은 시드로 실행하면 언제나 같은 100명이 뽑힙니다.
  **데이터를 더 내려받으면 모집단이 바뀌어** 결과도 달라지니 주의하세요.
- 시드, 모집단 수, 추출 목록은 `labeling\sample_selection_log.txt`에 남습니다.
- `sample_list`에 이미 표본이 기록돼 있으면 덮어쓰지 않습니다.
  다시 뽑으려면 `labeling\backup\`의 백업본으로 되돌리세요.

### 라벨링 프로그램 (`run_labeler.bat`)
실행하면 검은 창(콘솔)과 프로그램 창이 함께 뜹니다. **검은 창을 닫으면 프로그램도 꺼지니** 그대로 두세요.
오류가 나면 검은 창에 메시지가 남으니 그 내용을 복사해 알려 주세요.

| 동작 | 방법 |
|---|---|
| 환자 이동 | 왼쪽 목록 클릭, ← → 키, [이전 환자]/[다음 환자] 버튼 |
| 슬라이스 넘기기 | 마우스 휠, ↑ ↓ 키, 이미지 아래 슬라이더 |
| 밝기/대비 | 이미지 위에서 오른쪽 버튼 드래그(위아래=밝기, 좌우=대비), 슬라이더, [초기화] |
| 좌표 확인 | 마우스를 올리면 아래 상태줄에 원본 DICOM 픽셀 x, y 표시 |
| 확대 / 이동 | **Ctrl + 휠** = 마우스 위치 중심 확대(최대 8배), **Ctrl + 왼쪽 드래그** 또는 휠 버튼 드래그 = 이동, **0 키**·[화면 맞춤] = 원래 크기. 슬라이스를 넘겨도 확대 유지, 환자를 바꾸면 원래 크기 |
| 레벨 표시 | F2 키, [레벨 표시] 체크 — L1/L2~L5/S1 위치에 초록점, 글자는 오른쪽으로 떨어진 곳에 선으로 연결 (진한 연두 = 좌표를 찍은 슬라이스, 흐린 초록 = 다른 슬라이스) |
| 좌표 슬라이스로 | Home 키 |
| 조작법 보기 | F1 키, [조작법 보기] 버튼 (기본은 숨김) |

- 메모리 절약을 위해 **현재 환자의 Sagittal T2 영상만** 읽고, 환자를 바꾸면 이전 영상은 버립니다.
- 좌표는 `train_label_coordinates.csv`와 같은 원본 픽셀 기준입니다(x=가로, y=세로, 왼쪽 위가 0).
- 시리즈 중간의 instance 번호가 비어 있으면 상태줄에 "일부 누락" 경고가 뜹니다.
- 레벨 위치는 `train_label_coordinates.csv`에서 현재 환자의 Sagittal T2 좌표만 그때그때 읽어 표시합니다.
  환자를 열면 좌표를 찍은 슬라이스부터 보여 줍니다(좌표가 없으면 가운데 슬라이스).

### 판독 입력 (④)
오른쪽 **판독 입력** 칸에서 하거나 키보드로 합니다. 키는 **영문 입력 상태**에서 쓰세요(한/영 키). 버튼은 언제나 됩니다.

| 키 | 동작 |
|---|---|
| **N** | 정상 → labels 11칸 모두 Normal, 저장 후 다음 환자로 (`config.py`의 `AUTO_NEXT_AFTER_NORMAL`) |
| **S** / **A** | 쉬모를 결절 / 애매 병변 입력 시작 |
| **1~6** | 척추 L1, L2, L3, L4, L5, S1 |
| **U** / **D** | 상종판 / 하종판 (S1은 상종판만) |
| 영상 클릭 | 병변 기록 (원본 픽셀 x, y와 현재 instance_number) |
| **Esc** | 병변 입력 끝내기 |
| **Del** | '기록된 병변' 목록에서 고른 병변 삭제 |
| **Ctrl+S** | 저장 다시 시도 (엑셀이 열려 있어 실패했을 때) |

- 병변이 여러 개면 **척추 → 종판 → 클릭**을 반복합니다. 병변 1개 = `lesions` 시트 1행.
- `labels`: 병변이 있는 종판 칸만 Schmorl/Uncertain, 나머지는 Normal. 같은 종판에 S와 A가 함께 있으면 Schmorl.
- 애매(A) 병변은 `review_status`에 **Pending**을 넣고(칸이 비어 있을 때만), 클릭 위치를 표시한 슬라이스를
  `labeling\captures\{study_id}_{척추}{sup|inf}.png`로 저장해 `capture_file`에 적습니다.
- `sample_list`의 status가 함께 바뀝니다: 미완료 / 완료 / 애매포함.
- **입력할 때마다 바로 엑셀에 저장**합니다. 저장할 때마다 엑셀 파일을 새로 읽어 그 환자 칸만 고치므로,
  교수님이 노란 칸(review_status, final_label, note)에 적은 내용은 지워지지 않습니다.
  교수님 입력이 있는 병변은 프로그램에서 지울 수 없습니다(엑셀에서 직접 확인).
- 엑셀이 열려 있어 저장이 실패하면 경고창이 뜨고, 입력 내용은 기억했다가 엑셀을 닫은 뒤 Ctrl+S(또는 다음 입력) 때 저장합니다.
- 프로그램을 켤 때마다 `labeling\backup\`에 엑셀 백업(`session_start`)을 남깁니다.
- 병변 번호(lesion_id)는 `{study_id}_01`, `_02` …이며 지운 번호는 다시 쓰지 않습니다.

### 이어하기 (⑤)
- 판독 결과는 입력할 때마다 엑셀에 바로 저장됩니다(④).
- 마지막으로 보던 **환자와 슬라이스**를 `labeling\labeler_state.json`에 기록해 두었다가, 다시 실행하면 그 자리에서 시작합니다.
  환자를 바꿀 때마다 기록하므로 프로그램이 갑자기 꺼져도 마지막 환자부터 이어집니다.
- 이 파일을 지우면(또는 깨지면) 첫 '미완료' 환자부터 시작합니다.

### MicroDicom 자동 열기 (⑥)
- 환자를 바꾸고 **0.8초쯤 머물면** 그 환자의 Sagittal T2 폴더를 MicroDicom으로 엽니다.
  (← → 로 빠르게 넘기는 중에는 열지 않습니다)
- 이 프로그램이 열었던 이전 MicroDicom 창은 닫고 새로 엽니다. 직접 연 MicroDicom 창은 건드리지 않습니다.
- **M 키** 또는 [MicroDicom 열기] 버튼 = 지금 환자를 다시 열기, [자동 열기] 체크 = 자동 열기 끄고 켜기.
- 설정은 `config.py`:
  - `MICRODICOM_PATH` — 실행 파일 위치 (기본 `C:\Program Files\MicroDicom\mDicom.exe`)
  - `MICRODICOM_ARGS` — 넘기는 인자 (기본 `["-fd", "{folder}"]`, 공식 설명서의 `mDicom.exe -fd <폴더>`).
    MicroDicom 버전에 따라 폴더가 안 열리면 `["{folder}"]`로 바꿔 보세요.
  - `AUTO_OPEN_MICRODICOM`, `MICRODICOM_DELAY_MS`

## 파일 설명

| 파일 | 역할 |
|---|---|
| `config.py` | 경로, 시드, 표본 수, status 표기 같은 설정 |
| `dataset.py` | CSV와 DICOM 폴더를 읽는 공통 함수 |
| `excel_io.py` | 엑셀 머리글 색 판별, 백업, 안전 저장(열려 있으면 경고) |
| `check_setup.py` | ⓪ 점검 |
| `make_sample.py` | ① 100명 추출 |
| `remove_rater.py` | 엑셀 양식에서 rater 열 지우기 (한 번만) |
| `diagnose_excel.py` | 엑셀 파일 진단 (읽기 실패 원인, 정상 백업 찾기) |
| `workbook.py` | 엑셀 양식(sample_list, labels, lesions) 읽기/쓰기 |
| `imaging.py` | DICOM 시리즈 읽기, 밝기/대비 계산, 캡처 이미지 |
| `labeler.py` | 라벨링 프로그램 본체 |
| `microdicom.py` | MicroDicom 실행·이전 창 닫기 (⑥) |
