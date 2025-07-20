import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QVBoxLayout, QComboBox, QMessageBox
)
from PyQt5.QtCore import Qt
from datetime import datetime


class SmartExcelMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ادغام اطلاعات کاربران")
        self.setGeometry(100, 100, 700, 600)

        self.file_paths = []
        self.dataframes = {}

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        self.select_files_btn = QPushButton("انتخاب فایل‌ها")
        self.select_files_btn.setFixedSize(130, 30)
        self.select_files_btn.setStyleSheet("font-size: 14px; color: #030303; background-color: #9EC6F3; font-weight: bold;")
        self.select_files_btn.clicked.connect(self.select_files)
        layout.addWidget(self.select_files_btn, alignment=Qt.AlignCenter)

        self.files_label = QLabel("فایل بارگزاری نشده")
        self.files_label.setStyleSheet("font-size: 12px; color: #555; font-weight: bold;")
        layout.addWidget(self.files_label)

        self.ref_label = QLabel("انتخاب فایل مرجع:")
        self.ref_label.setStyleSheet("font-size: 12px; color: #333; font-weight: bold;")
        layout.addWidget(self.ref_label)

        self.ref_combo = QComboBox()
        self.ref_combo.setFixedSize(220, 25)
        self.ref_combo.setStyleSheet("font-size: 12px; font-weight: bold;")
        layout.addWidget(self.ref_combo, alignment=Qt.AlignRight)

        self.key_label = QLabel("نام ستون کلیدی (مثل کدملی):")
        self.key_label.setStyleSheet("font-size: 12px; color: #333; font-weight: bold;")
        layout.addWidget(self.key_label)

        self.key_input = QLineEdit("کد ملی دانشجو")
        self.key_input.setFixedSize(250, 35)
        self.key_input.setStyleSheet("font-size: 12px; padding: 5px; font-weight: bold;")
        layout.addWidget(self.key_input, alignment=Qt.AlignRight)

        self.merge_btn = QPushButton("ادغام و دریافت خروجی")
        self.merge_btn.setFixedSize(180, 40)
        self.merge_btn.setStyleSheet("font-size: 14px; color: white; background-color: #030303; font-weight: bold;")
        self.merge_btn.clicked.connect(self.merge_and_save)
        layout.addWidget(self.merge_btn, alignment=Qt.AlignCenter)

        layout.setSpacing(15)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        central_widget.setStyleSheet("background-color: #fdf6e3;")
        self.setCentralWidget(central_widget) 
        self.resize(600, 500)

    def select_files(self):
        self.files_label.setText("در حال بارگذاری فایل‌ها...")  # پیام در حال بارگذاری
        QApplication.processEvents()  

        paths, _ = QFileDialog.getOpenFileNames(self, "انتخاب فایل‌ها", "", "Excel Files (*.xlsx)")
        if paths:
            self.file_paths = paths
            self.dataframes.clear()

            for path in paths:
                df = pd.read_excel(path)
                self.dataframes[path] = df

            self.files_label.setText("\n".join([path.split("/")[-1] for path in paths]))
            self.ref_combo.clear()
            self.ref_combo.addItems([path.split("/")[-1] for path in paths])
        else:
            self.files_label.setText("هیچ فایلی انتخاب نشده")

    def merge_and_save(self):
        self.files_label.setText("در حال ادغام اطلاعات...")  # پیام در حال ادغام
        QApplication.processEvents()  # به‌روزرسانی فوری UI

        if not self.file_paths or len(self.file_paths) < 2:
            QMessageBox.warning(self, "خطا", "حداقل دو فایل انتخاب کنید.")
            self.files_label.setText("لطفاً فایل‌ها را دوباره انتخاب کنید.")
            return

        key_col = self.key_input.text().strip()
        if not key_col:
            QMessageBox.warning(self, "خطا", "ستون کلید (مثلاً کدملی) را وارد کنید.")
            return
        ref_filename = self.ref_combo.currentText()
        ref_path = next((p for p in self.file_paths if p.endswith(ref_filename)), None)
        if not ref_path:
            QMessageBox.warning(self, "خطا", "فایل مرجع انتخاب نشده.")
            return

        ref_df = self.dataframes[ref_path].copy()
        if key_col not in ref_df.columns:
            QMessageBox.warning(self, "خطا", f"ستون '{key_col}' در فایل مرجع یافت نشد.")
            return

        all_data = {row[key_col]: {} for _, row in ref_df.iterrows()}

        for path, df in self.dataframes.items():
            if path == ref_path or key_col not in df.columns:
                continue

            date_col = next((c for c in df.columns if 'تاریخ' in c or 'زمان' in c), None)

            for _, row in df.iterrows():
                key = row.get(key_col)
                if pd.isna(key) or key not in all_data:
                    continue

                for col in df.columns:
                    if col == key_col:
                        continue

                    val = row[col]
                    if pd.isna(val):
                        continue

                    if col not in all_data[key]:
                        all_data[key][col] = [(val, row.get(date_col) if date_col else None)]
                    else:
                        all_data[key][col].append((val, row.get(date_col) if date_col else None))

        # ساخت خروجی نهایی
        merged_rows = []
        for _, row in ref_df.iterrows():
            key = row[key_col]
            base = {key_col: key}
            info = all_data.get(key, {})

            for col, entries in info.items():
                try:
                    entries.sort(key=lambda x: pd.to_datetime(x[1]), reverse=True)
                except:
                    pass
                values = list(dict.fromkeys([str(e[0]) for e in entries]))
                base[col] = " / ".join(values)
            merged_rows.append(base)

        merged_df = pd.DataFrame(merged_rows)

        save_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", "merged.xlsx", "Excel Files (*.xlsx)")
        if save_path:
            merged_df.to_excel(save_path, index=False)
            QMessageBox.information(self, "انجام شد", f"فایل خروجی ذخیره شد: {save_path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SmartExcelMerger()
    window.show()
    sys.exit(app.exec_())
