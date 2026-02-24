import json
import shutil
import threading
from pathlib import Path

import numpy as np
import torch
import torch.nn as nn
import torch.optim as optim
from PIL import Image
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from torchvision import models, transforms

ROOT = Path.cwd()
WORKSPACE = ROOT / "workspace"
DATASET_DIR = WORKSPACE / "dataset"
EXPORT_DIR = WORKSPACE / "exports"
MODEL_DIR = WORKSPACE / "models"
SPLIT_FILE = DATASET_DIR / "splits.json"
META_FILE = MODEL_DIR / "metadata.json"
MODEL_FILE = MODEL_DIR / "classifier.pt"
ONNX_FILE = EXPORT_DIR / "classifier.onnx"


class TrainerSignals(QObject):
    progress = Signal(str)
    finished = Signal(bool, str)


class ClassifierGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Simple Classifier Studio")
        self.resize(1200, 760)

        for d in [DATASET_DIR, EXPORT_DIR, MODEL_DIR]:
            d.mkdir(parents=True, exist_ok=True)

        self.device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
        self.class_names = self._load_existing_classes()
        self.split_map = self._load_split_map()
        self.model = None

        self.transform_train = transforms.Compose([
            transforms.Resize((224, 224)),
            transforms.RandomHorizontalFlip(),
            transforms.ToTensor(),
        ])
        self.transform_eval = transforms.Compose([
            transforms.Resize((224, 224)),
            transforms.ToTensor(),
        ])

        self._build_ui()
        self.refresh_classes()

    def _build_ui(self):
        root = QWidget()
        layout = QGridLayout(root)

        data_group = QGroupBox("1) Dataset 구성")
        data_layout = QVBoxLayout(data_group)

        class_row = QHBoxLayout()
        self.class_input = QLineEdit()
        self.class_input.setPlaceholderText("새 클래스 이름 (예: cat)")
        self.add_class_btn = QPushButton("클래스 추가")
        self.add_class_btn.clicked.connect(self.add_class)
        class_row.addWidget(self.class_input)
        class_row.addWidget(self.add_class_btn)

        self.class_list = QListWidget()

        image_btn_row = QHBoxLayout()
        self.add_images_btn = QPushButton("선택 클래스에 이미지 추가")
        self.add_images_btn.clicked.connect(self.add_images_to_selected_class)
        self.preview_image_btn = QPushButton("선택 이미지 미리보기")
        self.preview_image_btn.clicked.connect(self.preview_selected_dataset_image)
        self.mark_train_btn = QPushButton("선택 이미지 → 학습용")
        self.mark_train_btn.clicked.connect(lambda: self.set_selected_image_split("train"))
        self.mark_eval_btn = QPushButton("선택 이미지 → 평가용")
        self.mark_eval_btn.clicked.connect(lambda: self.set_selected_image_split("eval"))
        image_btn_row.addWidget(self.add_images_btn)
        image_btn_row.addWidget(self.preview_image_btn)
        image_btn_row.addWidget(self.mark_train_btn)
        image_btn_row.addWidget(self.mark_eval_btn)

        self.dataset_images = QListWidget()

        data_layout.addLayout(class_row)
        data_layout.addWidget(QLabel("클래스"))
        data_layout.addWidget(self.class_list)
        data_layout.addWidget(QLabel("선택 클래스 이미지"))
        data_layout.addWidget(self.dataset_images)
        data_layout.addLayout(image_btn_row)

        train_group = QGroupBox("2) 학습")
        train_layout = QVBoxLayout(train_group)

        hp_row = QHBoxLayout()
        self.epoch_input = QLineEdit("3")
        self.batch_input = QLineEdit("8")
        self.lr_input = QLineEdit("0.001")
        hp_row.addWidget(QLabel("Epoch"))
        hp_row.addWidget(self.epoch_input)
        hp_row.addWidget(QLabel("Batch"))
        hp_row.addWidget(self.batch_input)
        hp_row.addWidget(QLabel("LR"))
        hp_row.addWidget(self.lr_input)

        self.train_btn = QPushButton("학습 시작")
        self.train_btn.clicked.connect(self.train_model)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        train_layout.addLayout(hp_row)
        train_layout.addWidget(self.train_btn)
        train_layout.addWidget(self.progress)
        train_layout.addWidget(self.log)

        infer_group = QGroupBox("3) 결과 보기 / 내보내기")
        infer_layout = QVBoxLayout(infer_group)

        self.infer_btn = QPushButton("테스트 이미지로 예측")
        self.infer_btn.clicked.connect(self.run_inference)
        self.result_label = QLabel("예측 결과: (없음)")

        self.image_preview = QLabel("이미지 미리보기")
        self.image_preview.setAlignment(Qt.AlignCenter)
        self.image_preview.setMinimumHeight(280)
        self.image_preview.setStyleSheet("border:1px solid #999;")

        self.export_btn = QPushButton("ONNX 내보내기 (C++ 연동용)")
        self.export_btn.clicked.connect(self.export_onnx)

        infer_layout.addWidget(self.infer_btn)
        infer_layout.addWidget(self.result_label)
        infer_layout.addWidget(self.image_preview)
        infer_layout.addWidget(self.export_btn)

        layout.addWidget(data_group, 0, 0)
        layout.addWidget(train_group, 0, 1)
        layout.addWidget(infer_group, 0, 2)

        self.class_list.itemSelectionChanged.connect(self.refresh_dataset_images)
        self.dataset_images.itemSelectionChanged.connect(self.preview_selected_dataset_image)

        self.setCentralWidget(root)

    def log_msg(self, msg: str):
        self.log.append(msg)

    def _load_existing_classes(self):
        if not DATASET_DIR.exists():
            return []
        return sorted([d.name for d in DATASET_DIR.iterdir() if d.is_dir()])

    def _load_split_map(self):
        if not SPLIT_FILE.exists():
            return {}
        try:
            with open(SPLIT_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {k: ("eval" if v == "eval" else "train") for k, v in data.items()}
        except Exception:
            return {}

    def _save_split_map(self):
        SPLIT_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(SPLIT_FILE, "w", encoding="utf-8") as f:
            json.dump(self.split_map, f, ensure_ascii=False, indent=2)

    def _key_for_image(self, cls: str, filename: str):
        return f"{cls}/{filename}"

    def refresh_classes(self):
        self.class_names = self._load_existing_classes()
        self.class_list.clear()
        for c in self.class_names:
            self.class_list.addItem(QListWidgetItem(c))
        self.refresh_dataset_images()

    def add_class(self):
        name = self.class_input.text().strip()
        if not name:
            QMessageBox.warning(self, "경고", "클래스 이름을 입력하세요.")
            return
        class_dir = DATASET_DIR / name
        class_dir.mkdir(parents=True, exist_ok=True)
        self.class_input.clear()
        self.refresh_classes()
        self.log_msg(f"클래스 추가: {name}")

    def selected_class(self):
        item = self.class_list.currentItem()
        if item is None:
            return None
        return item.text()

    def add_images_to_selected_class(self):
        cls = self.selected_class()
        if not cls:
            QMessageBox.warning(self, "경고", "먼저 클래스를 선택하세요.")
            return
        files, _ = QFileDialog.getOpenFileNames(self, "이미지 선택", str(ROOT), "Images (*.png *.jpg *.jpeg *.bmp)")
        if not files:
            return

        class_dir = DATASET_DIR / cls
        for fp in files:
            src = Path(fp)
            dst = class_dir / src.name
            if dst.exists():
                base, ext = src.stem, src.suffix
                idx = 1
                while (class_dir / f"{base}_{idx}{ext}").exists():
                    idx += 1
                dst = class_dir / f"{base}_{idx}{ext}"
            shutil.copy2(src, dst)
            self.split_map[self._key_for_image(cls, dst.name)] = "train"

        self._save_split_map()

        self.log_msg(f"{len(files)}개 이미지를 '{cls}' 클래스에 추가")
        self.refresh_dataset_images()

    def refresh_dataset_images(self):
        self.dataset_images.clear()
        cls = self.selected_class()
        if not cls:
            return
        for img in sorted((DATASET_DIR / cls).glob("*")):
            if img.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp"}:
                split = self.split_map.get(self._key_for_image(cls, img.name), "train")
                self.dataset_images.addItem(QListWidgetItem(f"[{split}] {img.name}"))

    def _selected_dataset_filename(self):
        item = self.dataset_images.currentItem()
        if item is None:
            return None
        text = item.text()
        if "] " in text:
            return text.split("] ", 1)[1]
        return text

    def set_selected_image_split(self, split: str):
        cls = self.selected_class()
        filename = self._selected_dataset_filename()
        if not cls or not filename:
            QMessageBox.warning(self, "경고", "클래스와 이미지를 먼저 선택하세요.")
            return

        self.split_map[self._key_for_image(cls, filename)] = split
        self._save_split_map()
        self.refresh_dataset_images()
        self.log_msg(f"{filename} → {split} 설정")

    def preview_selected_dataset_image(self):
        cls = self.selected_class()
        filename = self._selected_dataset_filename()
        if not cls or not filename:
            return
        img_path = DATASET_DIR / cls / filename
        self.show_preview(img_path)

    def show_preview(self, img_path: Path):
        pixmap = QPixmap(str(img_path))
        if pixmap.isNull():
            self.image_preview.setText("미리보기를 표시할 수 없습니다")
            return
        self.image_preview.setPixmap(pixmap.scaled(self.image_preview.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def build_model(self, num_classes: int, pretrained: bool = True):
        weights = models.ResNet18_Weights.DEFAULT if pretrained else None
        model = models.resnet18(weights=weights)
        model.fc = nn.Linear(model.fc.in_features, num_classes)
        return model.to(self.device)

    def create_dataset(self):
        train_images = []
        train_labels = []
        eval_images = []
        eval_labels = []
        classes = sorted([d.name for d in DATASET_DIR.iterdir() if d.is_dir()])
        for idx, cls in enumerate(classes):
            for img in (DATASET_DIR / cls).glob("*"):
                if img.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp"}:
                    split = self.split_map.get(self._key_for_image(cls, img.name), "train")
                    if split == "eval":
                        eval_images.append(img)
                        eval_labels.append(idx)
                    else:
                        train_images.append(img)
                        train_labels.append(idx)
        return train_images, train_labels, eval_images, eval_labels, classes

    def evaluate_model(self, model, images, labels, batch_size):
        if not images:
            return None, None

        model.eval()
        criterion = nn.CrossEntropyLoss()
        total_loss = 0.0
        correct = 0
        count = 0
        with torch.no_grad():
            for start in range(0, len(images), batch_size):
                batch_imgs = images[start : start + batch_size]
                batch_labels = labels[start : start + batch_size]
                x = torch.stack([
                    self.transform_eval(Image.open(p).convert("RGB")) for p in batch_imgs
                ]).to(self.device)
                y = torch.tensor(batch_labels, dtype=torch.long).to(self.device)
                out = model(x)
                loss = criterion(out, y)
                total_loss += loss.item() * len(batch_imgs)
                pred = out.argmax(dim=1)
                correct += (pred == y).sum().item()
                count += len(batch_imgs)

        return total_loss / count, correct / count

    def train_model(self):
        try:
            epochs = int(self.epoch_input.text().strip())
            batch_size = int(self.batch_input.text().strip())
            lr = float(self.lr_input.text().strip())
        except ValueError:
            QMessageBox.warning(self, "오류", "하이퍼파라미터 형식이 잘못되었습니다.")
            return

        train_images, train_labels, eval_images, eval_labels, classes = self.create_dataset()
        if len(classes) < 2:
            QMessageBox.warning(self, "오류", "최소 2개 클래스가 필요합니다.")
            return
        if len(train_images) < max(4, len(classes) * 2):
            QMessageBox.warning(self, "오류", "학습 이미지 수가 너무 적습니다.")
            return

        self.progress.setVisible(True)
        self.train_btn.setEnabled(False)
        self.log_msg(
            f"학습 시작 - train={len(train_images)}, eval={len(eval_images)}, classes={len(classes)}"
        )
        if not eval_images:
            self.log_msg("평가용 이미지가 없어 검증은 생략됩니다. (선택 이미지를 eval로 지정 가능)")

        signals = TrainerSignals()
        signals.progress.connect(self.log_msg)

        def on_finish(ok, msg):
            self.progress.setVisible(False)
            self.train_btn.setEnabled(True)
            if ok:
                QMessageBox.information(self, "완료", msg)
            else:
                QMessageBox.critical(self, "실패", msg)

        signals.finished.connect(on_finish)

        def worker():
            try:
                try:
                    model = self.build_model(len(classes), pretrained=True)
                except Exception as e:
                    signals.progress.emit(
                        f"사전학습 가중치 다운로드 실패로 랜덤 초기화로 진행합니다: {e}"
                    )
                    model = self.build_model(len(classes), pretrained=False)
                criterion = nn.CrossEntropyLoss()
                optimizer = optim.Adam(model.parameters(), lr=lr)

                idxs = np.arange(len(train_images))
                np.random.shuffle(idxs)

                for epoch in range(epochs):
                    model.train()
                    total_loss = 0.0
                    correct = 0
                    count = 0

                    for start in range(0, len(idxs), batch_size):
                        batch_ids = idxs[start : start + batch_size]
                        batch_x = []
                        batch_y = []
                        for i in batch_ids:
                            img = Image.open(train_images[i]).convert("RGB")
                            batch_x.append(self.transform_train(img))
                            batch_y.append(train_labels[i])

                        x = torch.stack(batch_x).to(self.device)
                        y = torch.tensor(batch_y, dtype=torch.long).to(self.device)

                        optimizer.zero_grad()
                        out = model(x)
                        loss = criterion(out, y)
                        loss.backward()
                        optimizer.step()

                        total_loss += loss.item() * len(batch_ids)
                        pred = out.argmax(dim=1)
                        correct += (pred == y).sum().item()
                        count += len(batch_ids)

                    signals.progress.emit(
                        f"Epoch {epoch+1}/{epochs} - loss={total_loss/count:.4f}, acc={correct/count:.4f}"
                    )

                    eval_loss, eval_acc = self.evaluate_model(model, eval_images, eval_labels, batch_size)
                    if eval_loss is not None:
                        signals.progress.emit(
                            f"Epoch {epoch+1}/{epochs} - val_loss={eval_loss:.4f}, val_acc={eval_acc:.4f}"
                        )

                MODEL_DIR.mkdir(parents=True, exist_ok=True)
                torch.save(model.state_dict(), MODEL_FILE)
                with open(META_FILE, "w", encoding="utf-8") as f:
                    json.dump({"classes": classes}, f, ensure_ascii=False, indent=2)

                self.model = model.eval()
                self.class_names = classes
                signals.finished.emit(True, f"학습 완료. 모델 저장: {MODEL_FILE}")
            except Exception as e:
                signals.finished.emit(False, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def load_trained_model(self):
        if self.model is not None:
            return True
        if not (MODEL_FILE.exists() and META_FILE.exists()):
            return False

        with open(META_FILE, "r", encoding="utf-8") as f:
            meta = json.load(f)
        classes = meta["classes"]
        model = self.build_model(len(classes), pretrained=False)
        state = torch.load(MODEL_FILE, map_location=self.device)
        model.load_state_dict(state)
        model.eval()
        self.model = model
        self.class_names = classes
        return True

    def run_inference(self):
        if not self.load_trained_model():
            QMessageBox.warning(self, "오류", "먼저 모델을 학습하세요.")
            return

        file, _ = QFileDialog.getOpenFileName(self, "테스트 이미지 선택", str(ROOT), "Images (*.png *.jpg *.jpeg *.bmp)")
        if not file:
            return

        path = Path(file)
        self.show_preview(path)

        img = Image.open(path).convert("RGB")
        x = self.transform_eval(img).unsqueeze(0).to(self.device)

        with torch.no_grad():
            out = self.model(x)
            prob = torch.softmax(out, dim=1)[0]
            idx = int(torch.argmax(prob).item())
            score = float(prob[idx].item())

        self.result_label.setText(f"예측 결과: {self.class_names[idx]} ({score:.2%})")

    def export_onnx(self):
        if not self.load_trained_model():
            QMessageBox.warning(self, "오류", "내보낼 모델이 없습니다. 먼저 학습하세요.")
            return

        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        dummy = torch.randn(1, 3, 224, 224).to(self.device)
        model_cpu = self.model.to("cpu")
        torch.onnx.export(
            model_cpu,
            dummy.cpu(),
            str(ONNX_FILE),
            input_names=["input"],
            output_names=["logits"],
            opset_version=12,
            dynamic_axes={"input": {0: "batch"}, "logits": {0: "batch"}},
        )

        with open(EXPORT_DIR / "labels.txt", "w", encoding="utf-8") as f:
            for c in self.class_names:
                f.write(c + "\n")

        QMessageBox.information(
            self,
            "내보내기 완료",
            f"ONNX: {ONNX_FILE}\n라벨: {EXPORT_DIR / 'labels.txt'}",
        )


def main():
    app = QApplication([])
    window = ClassifierGUI()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
