import sys
import os
import time
import threading
import datetime
import logging
import traceback
from openai import OpenAI

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLineEdit, QLabel, QListWidget, QProgressBar, 
    QFileDialog, QMessageBox, QComboBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QCheckBox, QMenu
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QThread
import keyring
from core_engine import TranslationEngine
from doc_processor import DocProcessor
from glossary_manager import GlossaryManager

# 配置日志
def setup_global_logging(base_dir):
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"{datetime.datetime.now().strftime('%Y%m%d')}.log")
    
    # 同时输出到文件和控制台
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

class TranslationThread(QThread):
    progress_signal = pyqtSignal(int, str, str)  # percent, status, eta
    file_finished_signal = pyqtSignal(str)
    file_error_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal(bool)

    def __init__(self, files, api_key, base_url, model, base_dir, processor, output_dir, direction):
        super().__init__()
        self.files = files
        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.base_dir = base_dir
        self.processor = processor
        self.output_dir = output_dir
        self.direction = direction
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def run(self):
        total = len(self.files)
        if total == 0:
            self.finished_signal.emit(True)
            return

        start_ts = time.time()

        try:
            # 关键修复：把引擎与术语库加载放到工作线程里，避免 GUI 主线程阻塞导致“未响应”
            engine = TranslationEngine(
                api_key=self.api_key,
                base_url=self.base_url,
                model=self.model,
                base_dir=self.base_dir,
                direction=self.direction,
                stop_event=self._stop_event,
            )

            for i, path in enumerate(self.files, start=1):
                if self._stop_event.is_set():
                    self.progress_signal.emit(int((i - 1) / total * 100), "已停止。", "--")
                    self.finished_signal.emit(False)
                    return

                status = f"[{i}/{total}] {os.path.basename(path)}"
                percent = int(i / total * 100)

                elapsed = time.time() - start_ts
                rate = elapsed / i if i else 0
                remaining = rate * (total - i)
                eta = f"{int(remaining)}s" if remaining >= 1 else "0s"

                self.progress_signal.emit(percent, status, eta)

                try:
                    out_path = self.processor.translate_file(path, self.output_dir, engine)
                    self.file_finished_signal.emit(out_path)
                except Exception as e:
                    self.file_error_signal.emit(path, str(e))

            self.progress_signal.emit(100, "完成。", "0s")
            self.finished_signal.emit(True)

        except Exception as e:
            self.file_error_signal.emit("初始化/运行错误", str(e))
            self.finished_signal.emit(False)


class ApiTestWorker(QThread):
    ok = pyqtSignal(str)
    fail = pyqtSignal(str)

    def __init__(self, api_key: str, base_url: str):
        super().__init__()
        self.api_key = api_key
        self.base_url = base_url

    def run(self):
        try:
            client = OpenAI(api_key=self.api_key, base_url=self.base_url)
            _ = client.models.list()
            self.ok.emit("API 测试成功。")
        except Exception as e:
            self.fail.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.base_dir = os.getcwd()
        setup_global_logging(self.base_dir)
        self.logger = logging.getLogger("MainWindow")
        
        self.glossary = GlossaryManager(self.base_dir)
        self.processor = DocProcessor()
        self.files_to_translate = []
        self.custom_output_dir = None
        self.trans_thread = None
        
        self.setWindowTitle("工厂官方文件双向翻译软件 v2.2 (稳定版)")
        self.setMinimumSize(1000, 750)
        self.init_ui()
        self.load_config()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Tab 1: Translation
        trans_tab = QWidget()
        self.tabs.addTab(trans_tab, "翻译任务")
        self.init_trans_ui(trans_tab)

        # Tab 2: Glossary Management
        glossary_tab = QWidget()
        self.tabs.addTab(glossary_tab, "术语库管理")
        self.init_glossary_ui(glossary_tab)

    def init_trans_ui(self, widget):
        layout = QVBoxLayout(widget)
        
        # API Key Section
        api_layout = QHBoxLayout()
        self.api_key_input = QLineEdit()
        self.api_key_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key_input.setPlaceholderText("输入 DeepSeek API Key")
        self.save_key_btn = QPushButton("保存 Key")
        self.save_key_btn.clicked.connect(self.save_api_key)
        self.test_conn_btn = QPushButton("测试连接")
        self.test_conn_btn.clicked.connect(self.test_connection)
        
        api_layout.addWidget(QLabel("API Key:"))
        api_layout.addWidget(self.api_key_input)
        api_layout.addWidget(self.save_key_btn)
        api_layout.addWidget(self.test_conn_btn)
        layout.addLayout(api_layout)

        # Direction & Output Folder
        folder_layout = QHBoxLayout()
        self.direction_combo = QComboBox()
        self.direction_combo.addItems(["中 -> 英 (ZH to EN)", "英 -> 中 (EN to ZH)"])
        self.import_btn = QPushButton("导入文件")
        self.import_btn.clicked.connect(self.import_files)
        self.set_out_btn = QPushButton("选择输出目录")
        self.set_out_btn.clicked.connect(self.select_output_dir)
        self.output_dir_label = QLabel("输出目录: (默认)")
        self.output_dir_label.setStyleSheet("color: gray;")
        
        folder_layout.addWidget(QLabel("方向:"))
        folder_layout.addWidget(self.direction_combo)
        folder_layout.addWidget(self.import_btn)
        folder_layout.addWidget(self.set_out_btn)
        folder_layout.addWidget(self.output_dir_label)
        layout.addLayout(folder_layout)

        # Lists
        list_layout = QHBoxLayout()
        self.input_list = QListWidget()
        self.input_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.input_list.customContextMenuRequested.connect(self.show_input_context_menu)
        
        self.output_list = QListWidget()
        list_layout.addWidget(self.input_list)
        list_layout.addWidget(self.output_list)
        layout.addLayout(list_layout)

        # Progress
        progress_info = QHBoxLayout()
        self.status_label = QLabel("就绪")
        self.eta_label = QLabel("ETA: --")
        progress_info.addWidget(self.status_label)
        progress_info.addStretch()
        progress_info.addWidget(self.eta_label)
        layout.addLayout(progress_info)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Controls
        ctrl_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始翻译")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setStyleSheet("background-color: #0078D4; color: white; font-weight: bold;")
        self.start_btn.clicked.connect(self.start_translation)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setFixedHeight(40)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_translation)
        ctrl_layout.addWidget(self.start_btn)
        ctrl_layout.addWidget(self.stop_btn)
        layout.addLayout(ctrl_layout)

    def init_glossary_ui(self, widget):
        layout = QVBoxLayout(widget)
        self.glossary_stats = QLabel(f"当前条目数: {self.glossary.data['meta']['count']}")
        layout.addWidget(self.glossary_stats)

        self.glossary_table = QTableWidget(0, 2)
        self.glossary_table.setHorizontalHeaderLabels(["中文 (ZH)", "英文 (EN)"])
        self.glossary_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.glossary_table)
        self.refresh_glossary_table()

        btn_layout = QHBoxLayout()
        self.import_glossary_btn = QPushButton("导入术语 (Excel/Word)")
        self.import_glossary_btn.clicked.connect(self.import_glossary)
        self.add_term_btn = QPushButton("新增条目")
        self.add_term_btn.clicked.connect(self.add_term_dialog)
        btn_layout.addWidget(self.import_glossary_btn)
        btn_layout.addWidget(self.add_term_btn)
        layout.addLayout(btn_layout)

    def show_input_context_menu(self, pos):
        menu = QMenu()
        remove_action = menu.addAction("移除选中文件")
        clear_action = menu.addAction("清空列表")
        action = menu.exec(self.input_list.mapToGlobal(pos))
        if action == remove_action:
            self.remove_selected_files()
        elif action == clear_action:
            self.clear_input_list()

    def remove_selected_files(self):
        indices = [self.input_list.row(item) for item in self.input_list.selectedItems()]
        for index in sorted(indices, reverse=True):
            self.input_list.takeItem(index)
            self.files_to_translate.pop(index)

    def clear_input_list(self):
        self.input_list.clear()
        self.files_to_translate = []

    def select_output_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.custom_output_dir = path
            self.output_dir_label.setText(f"输出目录: {path}")

    def test_connection(self):
        key = self.api_key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "提示", "请输入 API Key")
            return

        # 注意：不能在非 GUI 线程里弹窗或修改控件，否则会出现 QObject::setParent 跨线程错误并导致卡死/崩溃
        self.test_conn_btn.setEnabled(False)
        self.test_conn_btn.setText("测试中...")

        base_url = "https://api.deepseek.com"
        self._api_test_worker = ApiTestWorker(key, base_url)
        self._api_test_worker.ok.connect(self._on_api_test_ok)
        self._api_test_worker.fail.connect(self._on_api_test_fail)
        self._api_test_worker.finished.connect(self._on_api_test_done)
        self._api_test_worker.start()

    def _on_api_test_ok(self, msg: str):
        QMessageBox.information(self, "成功", msg)

    def _on_api_test_fail(self, err: str):
        QMessageBox.critical(self, "失败", f"API 测试失败: {err}")

    def _on_api_test_done(self):
        self.test_conn_btn.setEnabled(True)
        self.test_conn_btn.setText("测试连接")

    def save_api_key(self):
        key = self.api_key_input.text().strip()
        if key:
            try:
                keyring.set_password("FactoryTranslator", "deepseek_api_key", key)
                QMessageBox.information(self, "成功", "API Key 已加密保存。")
            except:
                # 备选方案：保存到本地加密文件（此处简化为普通文件，实际可增加加密逻辑）
                with open(".api_key", "w") as f: f.write(key)
                QMessageBox.information(self, "成功", "API Key 已保存到本地文件。")

    def load_config(self):
        key = keyring.get_password("FactoryTranslator", "deepseek_api_key")
        if not key and os.path.exists(".api_key"):
            with open(".api_key", "r") as f: key = f.read().strip()
        if key: self.api_key_input.setText(key)

    def import_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "Documents (*.docx *.txt)")
        for f in files:
            if f not in self.files_to_translate:
                self.files_to_translate.append(f)
                self.input_list.addItem(os.path.basename(f))

    def start_translation(self):
        key = self.api_key_input.text().strip()
        if not key:
            QMessageBox.warning(self, "提示", "请输入 API Key")
            return
        if not self.files_to_translate:
            QMessageBox.warning(self, "提示", "请先导入要翻译的文件。")
            return

        self.output_list.clear()
        self.set_ui_enabled(False)

        base_url = "https://api.deepseek.com"
        model = "deepseek-chat"

        output_dir = self.custom_output_dir if self.custom_output_dir else os.path.join(self.base_dir, "output_en" if self.direction_combo.currentIndex() == 0 else "output_zh")
        os.makedirs(output_dir, exist_ok=True)

        direction = "zh2en" if self.direction_combo.currentIndex() == 0 else "en2zh"

        self.processor = self.processor if hasattr(self, "processor") and self.processor else DocProcessor()

        self.trans_thread = TranslationThread(
            files=self.files_to_translate,
            api_key=key,
            base_url=base_url,
            model=model,
            base_dir=self.base_dir,
            processor=self.processor,
            output_dir=output_dir,
            direction=direction,
        )
        self.trans_thread.progress_signal.connect(self.on_progress)
        self.trans_thread.file_finished_signal.connect(self.on_file_finished)
        self.trans_thread.file_error_signal.connect(self.on_error)
        self.trans_thread.finished_signal.connect(self.on_finished)
        self.trans_thread.start()

    def stop_translation(self):
        if self.trans_thread:
            self.trans_thread.stop()
            self.status_label.setText("正在停止...")

    def set_ui_enabled(self, enabled):
        self.start_btn.setEnabled(enabled)
        self.stop_btn.setEnabled(not enabled)
        self.import_btn.setEnabled(enabled)
        self.test_conn_btn.setEnabled(enabled)
        self.tabs.setTabEnabled(1, enabled)

    def on_progress(self, val, status, eta):
        self.progress_bar.setValue(val)
        self.status_label.setText(status)
        self.eta_label.setText(f"ETA: {eta}")

    def on_file_finished(self, out_path: str):
        # 输出列表显示生成文件名，并可用于后续“打开文件夹/打开文件”等功能扩展
        self.output_list.addItem(os.path.basename(out_path))

    def on_error(self, filename, msg):
        QMessageBox.warning(self, "文件处理错误", f"文件 {filename} 处理失败:\n{msg}")

    def on_finished(self, ok: bool):
        self.set_ui_enabled(True)
        self.status_label.setText("任务结束")
        self.eta_label.setText("ETA: --")

        if ok:
            QMessageBox.information(self, "完成", f"任务处理结束。请在右侧输出列表查看结果。")
        else:
            QMessageBox.warning(self, "结束", "任务已停止或发生错误，未完全完成。")

    def refresh_glossary_table(self):
        terms = self.glossary.data["zh2en"]
        self.glossary_table.setRowCount(len(terms))
        for i, (zh, en) in enumerate(terms.items()):
            self.glossary_table.setItem(i, 0, QTableWidgetItem(zh))
            self.glossary_table.setItem(i, 1, QTableWidgetItem(en))
        self.glossary_stats.setText(f"当前条目数: {len(terms)}")

    def import_glossary(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择术语文件", "", "Files (*.xlsx *.xls *.docx)")
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in ['.xlsx', '.xls']: self.glossary.import_excel(f)
            elif ext == '.docx': self.glossary.import_docx(f)
        self.refresh_glossary_table()

    def add_term_dialog(self):
        row = self.glossary_table.rowCount()
        self.glossary_table.insertRow(row)
        self.glossary_table.cellChanged.connect(self.on_cell_changed)

    def on_cell_changed(self, row, col):
        zh = self.glossary_table.item(row, 0)
        en = self.glossary_table.item(row, 1)
        if zh and en:
            self.glossary.add_term(zh.text(), en.text())
            self.glossary.save()
            self.refresh_glossary_table()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
