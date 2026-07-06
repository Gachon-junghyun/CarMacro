# 다른 PC에서 exe로 실행하기 (파이썬 설치 불필요)

타겟 PC에 파이썬/pip가 없거나 3.9라 패키지 설치가 안 될 때, **exe 하나만 복사**하면 된다.
exe 안에 파이썬과 selenium 등 패키지가 통째로 들어 있다.

## 빌드 (패키지가 있는 PC에서 1회)

```bat
build_exe.bat
```
→ `dist\CarMacro-app.exe` 생성 (약 36MB). 파이썬 3.x + 인터넷 있는 PC에서 실행.
(내부적으로 `pip install pyinstaller` 후 `app.py` 를 onefile 로 묶는다. pandas/numpy 는 제외.)

## 타겟 PC 요건
- **Windows** (빌드한 것과 같은 계열, 보통 64bit)
- **Chrome 설치** (브라우저 자체는 필요)
- **인터넷** — 최초 실행 시 selenium 이 Chrome 버전에 맞는 chromedriver 를 자동 다운로드
- 파이썬/pip 는 **불필요**

## 타겟 PC에 복사할 것
| 파일 | 용도 |
|---|---|
| `CarMacro-app.exe` | 실행파일(자동입력 앱) |
| `run_chrome.bat` | 크롬을 디버그 포트 9222로 띄우기 |
| 신청서 xlsx (예: `real_MMDD\신청서_이름.xlsx`) | 입력할 데이터 |
| 첨부 PDF | 첨부파일 업로드용 |

## 실행 순서 (타겟 PC)
1. `run_chrome.bat` 실행 → 뜬 크롬에서 **ev.or.kr 로그인 → 지자체 선택 → 신청서 폼** 진입
2. `CarMacro-app.exe` 더블클릭 → GUI
3. **[엑셀 불러오기]** → 신청서 xlsx 선택 → **[시작]**
4. 첫 실행은 chromedriver 다운로드로 몇 초 걸릴 수 있음(인터넷 필요)
5. 🔴 빨간 로그 확인, 저장/제출/암호는 직접

## 자주 나는 문제
- **"크롬 연결 실패 (포트 9222)"** → run_chrome.bat 로 띄운 크롬이 떠 있는지, 다른 크롬은 다 끈 상태인지 확인.
- **chromedriver 관련 오류** → 인터넷 연결 확인(최초 1회 다운로드 필요). 사내망이면 프록시/차단 확인.
- **표형식(CSV) 못 읽음** → 이 exe 는 pandas 를 뺐다. **세로형 양식**(신청서_이름.xlsx)을 쓸 것.
- 백신이 낯선 exe 를 막으면 예외 등록.

## 참고
- 데이터셋(xlsx) 만들기는 `make_gui.py`(소스) 또는 별도로 exe 를 뽑아 쓸 수 있다.
  make_gui 는 selenium 이 필요 없어 더 가볍게 빌드된다:
  `python -m PyInstaller --onefile --windowed --name CarMacro-maker --exclude-module pandas --exclude-module numpy --noconfirm make_gui.py`
