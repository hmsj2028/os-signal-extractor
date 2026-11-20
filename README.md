# OS Signal Extractor

Simulink(.slx) 모델 안의 기능 블록마다 **어떤 신호를 입력받아 어떤 신호를 내보내는지**를 자동으로 추적해 엑셀 표로 정리하는 MATLAB 도구입니다.

## 왜 만들었나

차량 제어 소프트웨어 모델에는 CAN 신호를 처리하는 기능 블록이 수백 개 들어 있습니다. 블록 하나의 입력이 실제로 어디서 오는지 알려면 Goto/From 태그, 게이트 블록, 여러 층의 서브시스템을 사람이 직접 따라가야 했습니다. 모델 하나에 몇 시간이 걸렸고, 놓치는 신호도 생겼습니다.

이 도구는 같은 일을 몇 초에서 몇 분 안에 끝냅니다. 직접 만들어 실무에서 쓰면서 v0.1부터 v0.8까지 고쳐 왔습니다. 실제 모델 2개(leaf 블록 294개)를 19초에 처리했습니다.

## 하는 일

1. 모델에서 이름이 `Swc_`로 시작하는 서브시스템(소프트웨어 컴포넌트)을 계층 깊이와 상관없이 모두 찾습니다. 이름에 `_sch`가 들어간 스케줄러는 뺍니다.
2. 그 안에서 내부 입력 포트 이름이 `PhyVal`인 마스킹 블록을 **leaf 기능 블록**으로 봅니다.
3. leaf 블록의 `PhyVal`(그리고 `ErrFlag`) 포트에서 신호를 **거꾸로 추적**합니다.
   - Gain, Data Type Conversion 같은 블록은 그냥 지나가고, SWC 경계 입력에 닿으면 멈춥니다.
   - **Goto/From**은 같은 태그의 Goto를 찾아 계속 따라갑니다. 같은 태그가 여러 개면 계층상 가장 가까운 것을 고릅니다.
   - **OR, Switch 같은 게이트 블록**을 만나면 들어오는 신호를 모두 `a, b, c` 형태로 나열합니다.
   - **Ground**와 **Constant**는 따로 표시합니다.
   - 깊이 제한(기본 30)을 두어 순환 참조가 있어도 멈추지 않게 했습니다.
4. 같은 Output이 여러 번 나오면 하나만 남기고, Input이 서로 다르면 로그에 경고를 남깁니다.
5. 모델별 시트와 요약(Summary) 시트가 담긴 엑셀, 그리고 **경고와 오류만 모은 로그**를 만듭니다.

이름 규칙(`Swc_`, `PhyVal`, `ErrFlag`, `_sch`)과 게이트로 볼 블록 종류는 코드 맨 위의 `CFG` 설정값에서 바꿀 수 있습니다.

## 예제로 바로 해 보기

`examples/make_example_project.m`은 추적 경우별 블록이 들어 있는 가짜 예제 모델을 만듭니다.

```matlab
cd examples
root = make_example_project();          % examples/example_project 생성
addpath(fullfile(root, 'os', 'os_demo'));
addpath('../data');
load_system('os_demo');
Os_signal_extractor({'os_demo/Swc_Demo'}, pwd)
```

결과는 아래와 같습니다. (`examples/sample_result.xlsx`에 같은 파일이 있습니다.)

| Input | ErrFlagInput | Output | Note |
|---|---|---|---|
| DefaultLv | | Out_DefaultLv | constant(os_demo_DefaultLevel) |
| DoorSts_FL, DoorSts_FR | | Out_DoorOpen | 다중입력 2건 |
| | | Out_Reserved | ground(0) |
| SpdA, SelCond, SpdB | | Out_SelSpd | 다중입력 3건 |
| VehSpd | SpdErr | Out_VehSpd | |

- `Out_VehSpd`: Gain을 지나 SWC 입력 `VehSpd`까지 추적했고, `ErrFlag`는 Goto/From을 따라가 `SpdErr`를 찾았습니다.
- 스케줄러인 `Swc_Demo_sch`는 대상 목록에서 빠집니다.

## 실제 프로젝트에서 쓰기 (UI 모드)

`run_extractor.bat`을 더블클릭하면 MATLAB이 열리고 다음 순서로 진행됩니다.

1. 상위 폴더를 고릅니다. 하위 전체가 MATLAB 경로에 추가됩니다.
2. `os` 폴더와 그 아래 모델 폴더(`os_xxx`)를 체크리스트에서 고릅니다.
3. `rte` 폴더의 `.mat` 파일과 모델별 `<모델명>_param.m`을 먼저 불러온 뒤 모델을 엽니다.
4. 모델마다 대상 SWC를 체크리스트에서 고릅니다. 기본값은 전체 선택입니다.
5. 끝나면 `results/result_<날짜시간>.xlsx`와 `logs/log_<날짜시간>.txt`가 생깁니다.

예상하는 폴더 구조는 다음과 같습니다.

```
<상위 폴더>/
  rte/*.mat
  os/os_xxx/os_xxx.slx
  os/os_xxx/os_xxx_param.m
```

## 폴더 구성

```
data/Os_signal_extractor.m          도구 본체 (약 1,300줄)
run_extractor.bat                   더블클릭 실행기 (data 폴더의 .p → .m 순서로 찾아 실행)
examples/make_example_project.m     예제 모델 생성 스크립트
examples/sample_result.xlsx         예제 실행 결과
```

## 버전 기록

- **v0.1**: 단일 .slx 선택 → Swc_ 서브시스템 자동 탐색 → PhyVal 포트 기준 Input/Output 추출 → 모델별 엑셀 저장
- **v0.4**: 상위 폴더 하나로 os/rte 폴더 자동 인식, 여러 모델 일괄 처리, Goto/From을 SWC 경계까지 끝까지 추적, 결과 통합 저장
- **v0.6**: 마스크·라이브러리 안의 Goto까지 검색, 같은 태그가 여러 개일 때 가장 가까운 Goto 선택, 단계별 진행률 표시
- **v0.8**: 여러 모델을 고르면 로드를 먼저 끝내고 선택 창을 하나씩 차례로 띄움, 밑줄이 든 이름이 아래첨자로 보이던 표시 오류 수정

## 알려진 한계

- 입력이 여러 개인데 게이트로 등록되지 않은 블록(Sum, Product 등)은 첫 번째 입력만 따라갑니다. 필요하면 `CFG.GateBlockTypes`에 추가하면 됩니다.
- 엑셀 헤더 서식(색, 굵게)은 Windows와 Excel이 있을 때만 적용됩니다. 표 자체는 Excel이 없어도 만들어집니다.

## 환경

MATLAB, Simulink (R2026a에서 확인)
