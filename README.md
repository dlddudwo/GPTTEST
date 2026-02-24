# Simple Classifier GUI (PySide6 + PyTorch)

빈 프로젝트를 **이미지 분류 GUI 학습툴**로 구성했습니다.

원하는 목표:
- 사용자가 GUI에 이미지 추가
- 분류(classify) 모델 학습
- 예측 결과/이미지 확인
- 모델을 외부(C++)에서 쓰기 위해 내보내기

위 흐름을 전부 포함합니다.

---

## 1) 설치

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 2) 실행

```bash
python src/main.py
```

## 3) 사용 방법

1. **클래스 추가** (예: `cat`, `dog`)
2. 클래스 선택 후 **이미지 추가**
3. 선택 이미지를 **학습용(train)** 또는 **평가용(eval)** 으로 지정
4. Epoch/Batch/LR 설정 후 **학습 시작** (평가용 이미지가 있으면 에폭별 검증 로그 출력)
5. **테스트 이미지로 예측**
6. **ONNX 내보내기** 버튼으로 C++ 연동 파일 생성

## 4) 출력 파일

- 학습 모델: `workspace/models/classifier.pt`
- 클래스 메타: `workspace/models/metadata.json`
- ONNX 모델: `workspace/exports/classifier.onnx`
- 라벨 파일: `workspace/exports/labels.txt`

## 5) C++ 연동

`cpp_example/README.md` 참고.

---

## 참고

- 기본 백본은 `resnet18` 입니다.
- 데이터가 아주 적으면 성능이 낮을 수 있습니다.
- GPU가 있으면 자동 사용, 없으면 CPU로 동작합니다.
