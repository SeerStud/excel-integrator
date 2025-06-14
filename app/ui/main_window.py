import os
import logging
from PySide6 import QtWidgets
from app.config import load_rules
from app.ui.dialogs import RulesManagerDialog
from app.processing.reader import process_files
from app.processing.writer import save_result

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Интегратор Excel")
        self.rules = load_rules()
        self.files = []
        self.log = []
        self._init_ui()

    def _init_ui(self):
        cw = QtWidgets.QWidget()
        v  = QtWidgets.QVBoxLayout(cw)

        self.listw = QtWidgets.QListWidget()
        v.addWidget(self.listw)

        h = QtWidgets.QHBoxLayout()
        b_add = QtWidgets.QPushButton("Добавить")
        b_add.clicked.connect(self.add_files)
        b_del = QtWidgets.QPushButton("Удалить")
        b_del.clicked.connect(self.remove_files)
        h.addWidget(b_add)
        h.addWidget(b_del)
        v.addLayout(h)

        b_rules = QtWidgets.QPushButton("Правила")
        b_rules.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        b_rules.clicked.connect(self.manage_rules)
        v.addWidget(b_rules)

        b_merge = QtWidgets.QPushButton("Объединить")
        b_merge.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        b_merge.clicked.connect(self.merge_all)
        v.addWidget(b_merge)

        self.setCentralWidget(cw)

    def add_files(self):
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Выберите файлы", "", "Excel (*.xls *.xlsx)"
        )
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listw.addItem(p)

    def remove_files(self):
        for it in self.listw.selectedItems():
            p = it.text()
            self.files.remove(p)
            self.listw.takeItem(self.listw.row(it))

    def manage_rules(self):
        dlg = RulesManagerDialog(self, self.rules)
        dlg.exec()
        self.rules = load_rules()

    def merge_all(self):
        if not self.files:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Добавьте файлы для объединения")
            return
        try:
            result, self.log = process_files(self.files, self.rules)
            save_result(result, self.log)
        except Exception as e:
            logging.exception("Ошибка при объединении файлов")
            QtWidgets.QMessageBox.critical(
                self,
                "Ошибка",
                f"{e}\n\nПодробности см. в файле app.log"
            )
