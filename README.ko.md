# svg2inkppt

**SVG 로고를 넣으면, 슬라이드 쇼에서 획이 하나씩 그려지는 PowerPoint 파일이 나옵니다.**

[English README](README.md)

PowerPoint에는 도형의 path를 따라 선이 그려지는 애니메이션이 없습니다. 대신 **잉크**에는 있습니다.
펜으로 그린 획을 다시 그려주는 *재생(Replay)* 효과입니다. 이 도구는 SVG의 path를 PowerPoint
잉크(InkML)로 바꿔 슬라이드에 심고, 원하는 시간과 이징으로 재생 애니메이션을 붙입니다.
결과물은 평범한 `.pptx`와 평범한 애니메이션입니다. 다른 효과와 순서를 섞을 수 있고, 다른 덱에
복사해도 되고, PowerPoint만 있으면 어디서든 재생됩니다. 애드인도 매크로도 필요 없습니다.

```
python svg2inkppt.py logo.svg --duration 3 --ease-in 0.2 --ease-out 0.3
```

GUI를 쓰면 SVG를 드롭하고, path 순서를 드래그로 바꾸고, 슬라이더를 조절한 뒤 완료만 누르면 됩니다.

## 기능

- **SVG 도형 전부 지원**: `path`, `rect`, `circle`, `ellipse`, `line`, `polyline`, `polygon`,
  변환이 걸린 그룹. 일러스트레이터에서 내보낸 SVG가 그대로 됩니다.
- **끊긴 path** (M 명령 여러 개, 합쳐진 패스)는 획을 나눠 순서대로 그립니다.
- **선 색과 굵기**는 SVG 값을 따릅니다. 면만 있는 도형은 외곽선으로 그립니다.
- **일정한 펜 속도**: 모든 획을 호 길이 기준으로 다시 샘플링해서 긴 획과 짧은 획의 속도가 같습니다.
- **부드럽게 시작 / 끝** (ease-in / ease-out)을 애니메이션에 기록합니다.
- **그리기 순서 조절**: GUI에서 path 순서를 드래그로 바꾸고 획 방향을 뒤집을 수 있습니다.
  코드에서는 순서 목록을 넘기면 됩니다.
- **분리 모드**: SVG path마다 잉크 개체를 따로 만들어 각각 애니메이션 노드를 갖게 합니다.
  PowerPoint 애니메이션 창에서 개별로 시간을 조정할 수 있습니다.
- **템플릿 지원**: 기존 덱 마지막에 슬라이드를 추가해 테마를 그대로 씁니다.
- **대체 그림**: 잉크를 못 그리는 앱(Keynote, 구버전 Office, Google Slides)에서는 정지 PNG가 보입니다.

## 필요한 것

- Python 3.10 이상
- 애니메이션 *재생*에는 PowerPoint 2016 이상 (Microsoft 365 권장). PowerPoint for Mac 365에서 검증했습니다.
- GUI 실행 파일은 macOS용입니다. CLI와 코어는 플랫폼을 가리지 않습니다.

## 설치

```bash
git clone https://github.com/<you>/svg2inkppt.git
cd svg2inkppt
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## CLI

```bash
.venv/bin/python svg2inkppt.py logo.svg                      # -> logo.pptx
.venv/bin/python svg2inkppt.py logo.svg -o out.pptx --width-cm 12 --duration 3
.venv/bin/python svg2inkppt.py logo.svg --template deck.pptx # deck.pptx 마지막에 슬라이드 추가
```

| 옵션 | 의미 | 기본값 |
| --- | --- | --- |
| `-o, --output` | 출력 `.pptx` | `<svg 이름>.pptx` |
| `--template` | 슬라이드를 추가할 기존 덱 | 새 16:9 빈 덱 |
| `--width-cm`, `--height-cm` | 슬라이드 위 로고 크기 (비율 유지) | 슬라이드 폭의 60 %, 세로 최대 80 % |
| `--duration` | 총 그리는 시간(초) | `2` |
| `--ease-in`, `--ease-out` | 전체 시간 중 가속 / 감속 구간 비율, 0~1 | `0` (선형) |
| `--ease-mode` | `keyframes`는 이징 곡선을 키프레임 61개로 기록, `attr`는 `accel`/`decel` 속성 사용 | `keyframes` |
| `--stroke-width` | 펜 굵기 cm 강제 | SVG `stroke-width` 비례 |
| `--color` | 모든 획을 한 색으로 (`#RRGGBB`) | SVG 색 |
| `--split path` | SVG path마다 잉크 개체를 나눠 순서대로 재생 | 잉크 개체 하나 |
| `--step` | 샘플 간격 (SVG 단위, 작을수록 점이 많음) | 긴 변 / 600 |
| `--timing` | 잉크에 기록할 점별 시각 정보: `none`, `offset`, `channel` | `channel` |
| `--keep-png` | 대체 PNG를 출력 옆에 남김 | 끔 |

Python에서:

```python
from pathlib import Path
import svg2inkppt as core

opt = core.ConvertOptions(duration=3, ease_in=0.2, ease_out=0.3,
                          path_order=[2, 0, 1], reversed_paths={1})
core.convert(Path("logo.svg"), Path("logo.pptx"), opt)
```

## GUI (macOS)

`SVG2PPT.app`을 더블클릭합니다. 이 앱은 자기 위치를 기준으로 `.venv/bin/python app.py`를
실행하므로 프로젝트 폴더 안에 있어야 합니다. 터미널에서는 `.venv/bin/python app.py`.

1. 창에 SVG를 드롭합니다 (또는 클릭해서 선택).
2. 미리보기에 path마다 시작점에 순서 번호, 끝점에 점이 찍힙니다.
3. **그리기 순서** 목록에서 항목을 드래그해 순서를 바꿉니다. 하나를 고르면 그것만 강조됩니다.
   *방향 반전*은 path의 획 방향을 뒤집고, *원래대로*는 문서 순서로 되돌립니다.
4. 시간, ease-in / ease-out (작은 그래프가 곡선을 보여줍니다), 크기, 펜 굵기, 색, 분리 모드,
   템플릿 덱을 조절합니다.
5. **완료**를 누르고 저장 위치를 고르면 PowerPoint로 열립니다.

오류 로그는 `/tmp/svg2ppt.log`에 남습니다.

## 동작 원리

pptx는 펜 획을 InkML 파트(`ppt/ink/inkN.xml`)로 저장하고, 슬라이드에서 `p:contentPart`
요소로 참조합니다. 재생 효과는 `drawProgress` 속성을 0에서 1로 올리는 평범한 타이밍 노드
(`presetID="63"`)입니다. 변환기는 이렇게 동작합니다.

1. `svgelements`로 SVG를 읽고 변환을 펼친 뒤 모든 도형을 path 세그먼트로 바꿉니다.
2. path를 M 명령마다 나누고, 세그먼트를 호 길이 기준으로 점으로 샘플링합니다.
3. 점을 InkML trace로 씁니다 (1/1000 cm 단위, 1차 차분 인코딩). 색과 굵기 조합마다 붓 하나.
4. `python-pptx` 슬라이드에 잉크 파트, 대체 PNG, 타이밍 트리를 추가합니다.

실험으로 알아낸, 틀리기 쉬운 것들:

- PowerPoint는 재생 시간을 **점 개수** 비례로 나눕니다. 그래서 모든 획을 길이 기준으로 균일하게
  샘플링해야 합니다. 점 2개짜리 직선은 아주 느리게 그려집니다.
- 붓 속성 `fitToCurve`는 꺼야 합니다. 점이 적은 다각형이 물방울로 뭉개지고, 그리는 도중 선이 밖으로 튑니다.
- 재생 효과는 `accel`/`decel` 속성과 다중 키프레임 `tavLst`를 모두 존중합니다.
  PowerPoint UI에는 이징 조절이 노출되지 않을 뿐입니다.
- trace의 `timeOffset`과 `T` 채널은 받아들이지만 필수는 아닙니다.

## 한계

- 잉크는 선만 표현합니다. 면 채우기, 그라데이션, 점선, 화살표 끝은 옮겨지지 않습니다.
  면만 있는 도형은 외곽선으로 그립니다.
- 이징은 PowerPoint 안에서 나중에 바꿀 수 없습니다. 값을 바꿔 다시 생성해야 합니다.
- 펜 굵기와 색은 획 단위이고 끝은 둥급니다.
- 잉크 개체를 다른 덱으로 복사하면 애니메이션도 따라가지만, 사용자 정의 이징이 붙여넣기 뒤에도
  유지되는지는 검증하지 못했습니다. `--template`으로 대상 덱에 직접 생성하는 것이 안전합니다.

## 구성

```
svg2inkppt.py    코어 + CLI
app.py           Tkinter GUI
SVG2PPT.app      macOS 실행 번들 (셸 스크립트만 있고 Python은 들어있지 않음)
examples/        예제 SVG
internal/        개발용 덱과 테스트 출력물 (git 무시)
```

## 기여

이슈와 PR 환영합니다. 특히 아직 테스트하지 못한 Windows용 PowerPoint와 웹용 PowerPoint에서의
결과를 알려주시면 큰 도움이 됩니다.
