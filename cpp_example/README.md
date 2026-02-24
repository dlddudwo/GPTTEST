# C++ 연동 예시 (ONNX Runtime)

GUI에서 `ONNX 내보내기`를 누르면 `workspace/exports/classifier.onnx` 와 `labels.txt`가 생성됩니다.

이 파일들을 C++ 프로젝트로 복사한 뒤 ONNX Runtime C++ API로 로딩하면 됩니다.

핵심 흐름:
1. 이미지 전처리 (224x224, RGB, float32, [0,1])
2. `classifier.onnx` 추론
3. 출력 `logits`에서 argmax
4. `labels.txt` 인덱스로 클래스명 매핑

빌드는 ONNX Runtime 설치 방식(OS/패키지 매니저)에 따라 달라지므로, 이 저장소에서는 코드 골격만 제공합니다.
