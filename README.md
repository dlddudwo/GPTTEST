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
3. 필요 없는 샘플은 **선택 이미지 삭제**로 정리
4. 선택 이미지를 **학습용(train)** 또는 **평가용(eval)** 으로 지정
5. Epoch/Batch/LR + **Model 콤보박스(ResNet18/MobileNetV3/EfficientNet/ConvNeXt)** 선택 후 학습
   (평가용 이미지가 있으면 에폭별 검증 로그 출력)
6. 오른쪽 미리보기에서 **마우스 휠 확대/축소 + 드래그 이동**으로 이미지 자세히 확인
7. **Heatmap 보기 체크박스**를 켜고 예측하면 모델 중요영역(heatmap) 오버레이 확인
8. 학습 후 **평가 Confusion Matrix 보기** 버튼으로 eval 기준 confusion matrix 확인
9. **테스트 이미지로 예측** (학습 모델이 없으면 선택한 백본의 사전학습 모델로 판정)
10. 데이터셋 목록에서 여러 이미지를 선택한 뒤 **선택 이미지 판정(학습 없이 가능)** 으로 일괄 판정
11. **ONNX 내보내기** 버튼으로 C++ 연동 파일 생성

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
- 첫 학습 시 사전학습 가중치 다운로드가 실패(SSL/사내망)하면 자동으로 랜덤 초기화 학습으로 fallback 됩니다.
