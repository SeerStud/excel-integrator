from PySide6 import QtWidgets, QtCore
from PySide6.QtWidgets import QMessageBox
from app.config import save_rules

_original_question = QMessageBox.question
def _localized_question(parent, title, text,
                        buttons=QMessageBox.Yes | QMessageBox.No,
                        defaultButton=QMessageBox.No):
    mb = QMessageBox(parent)
    mb.setWindowTitle(title)
    mb.setText(text)
    mb.setStandardButtons(buttons)
    mb.setDefaultButton(defaultButton)
    btn_yes = mb.button(QMessageBox.Yes)
    if btn_yes: btn_yes.setText("Да")
    btn_no = mb.button(QMessageBox.No)
    if btn_no:  btn_no.setText("Нет")
    return mb.exec()
QMessageBox.question = _localized_question

def localize_buttonbox(bb: QtWidgets.QDialogButtonBox):
    btn_save = bb.button(QtWidgets.QDialogButtonBox.Save)
    if btn_save: btn_save.setText("Сохранить")
    btn_cancel = bb.button(QtWidgets.QDialogButtonBox.Cancel)
    if btn_cancel: btn_cancel.setText("Отмена")


class ColumnsDialog(QtWidgets.QDialog):
    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.cfg = cfg
        self.setWindowTitle("Столбцы")
        self.resize(400, 500)
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels([
            "Целевое", "Синонимы (через запятую)", "Не объединять"
        ])
        for rule in self.cfg.get("rules", []):
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0,
                QtWidgets.QTableWidgetItem(rule["target"]))
            self.table.setItem(r, 1,
                QtWidgets.QTableWidgetItem(",".join(rule.get("synonyms", []))))
            chk = QtWidgets.QTableWidgetItem()
            chk.setCheckState(
                QtCore.Qt.Checked if rule.get("no_merge", False)
                                 else QtCore.Qt.Unchecked
            )
            self.table.setItem(r, 2, chk)
        v.addWidget(self.table)

        btns = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить правило")
        self.btn_add.clicked.connect(self.add_rule)
        self.btn_del = QtWidgets.QPushButton("Удалить правило")
        self.btn_del.clicked.connect(self.del_rule)
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_del)
        v.addLayout(btns)

        self.chk_disable = QtWidgets.QCheckBox("Отключить")
        self.chk_disable.setChecked(not self.cfg.get("enabled", True))
        v.addWidget(self.chk_disable)
        self.chk_disable.stateChanged.connect(self._update_ui_state)

        self.grp_settings = QtWidgets.QGroupBox("Настройки NER")
        fl = QtWidgets.QFormLayout(self.grp_settings)

        self.spin_thr = QtWidgets.QSpinBox()
        self.spin_thr.setRange(0, 100)
        self.spin_thr.setValue(self.cfg.get("threshold", 80))
        fl.addRow("Порог схожести заголовков (0–100):", self.spin_thr)

        self.chk_auto = QtWidgets.QCheckBox("Объединять автоматически")
        self.chk_auto.setChecked(self.cfg.get("auto_merge", True))
        fl.addRow(self.chk_auto)

        self.chk_content = QtWidgets.QCheckBox("Анализ содержимого")
        self.chk_content.setChecked(self.cfg.get("use_content", False))
        fl.addRow(self.chk_content)

        self.spin_content_rows = QtWidgets.QSpinBox()
        self.spin_content_rows.setRange(1, 1000)
        self.spin_content_rows.setValue(self.cfg.get("content_rows", 10))
        fl.addRow("Число строк для анализа содержимого:", self.spin_content_rows)

        self.dsb_alpha = QtWidgets.QDoubleSpinBox()
        self.dsb_alpha.setRange(0.0, 1.0)
        self.dsb_alpha.setSingleStep(0.1)
        self.dsb_alpha.setValue(self.cfg.get("header_weight", 0.6))
        fl.addRow("Вес влияния заголовков (0-1):",
                  self.dsb_alpha)

        v.addWidget(self.grp_settings)

        self.chk_content.stateChanged.connect(self._update_content_state)
        self._update_content_state()

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        self._update_ui_state()
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())

    def _update_ui_state(self):
        disabled = self.chk_disable.isChecked()
        self.table.setDisabled(disabled)
        self.btn_add.setDisabled(disabled)
        self.btn_del.setDisabled(disabled)
        self.grp_settings.setDisabled(disabled)

    def _update_content_state(self):
        content_ok = self.chk_content.isChecked()
        self.spin_content_rows.setDisabled(not content_ok)
        self.dsb_alpha.setDisabled(not content_ok)

    def accept(self):
        self.cfg["enabled"] = not self.chk_disable.isChecked()

        rules = []
        for i in range(self.table.rowCount()):
            tgt = self.table.item(i, 0).text().strip()
            syn = [
                s.strip()
                for s in self.table.item(i, 1).text().split(",")
                if s.strip()
            ]
            no  = self.table.item(i, 2).checkState() == QtCore.Qt.Checked
            rules.append({
                "target": tgt,
                "synonyms": syn,
                "no_merge": no
            })
        self.cfg["rules"]        = rules
        self.cfg["threshold"]    = self.spin_thr.value()
        self.cfg["auto_merge"]   = self.chk_auto.isChecked()
        self.cfg["use_content"]  = self.chk_content.isChecked()
        self.cfg["content_rows"] = self.spin_content_rows.value()
        self.cfg["header_weight"]= self.dsb_alpha.value()

        super().accept()


class UnitsDialog(QtWidgets.QDialog):
    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.setWindowTitle("Единицы измерения")
        self.resize(350, 300)
        self.cfg = cfg
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(
            ["Колонка", "Целевая ед.", "Факторы"]
        )
        for rule in self.cfg.get("rules", []):
            r = self.table.rowCount(); self.table.insertRow(r)
            self.table.setItem(r, 0,
                QtWidgets.QTableWidgetItem(rule["column"]))
            self.table.setItem(r, 1,
                QtWidgets.QTableWidgetItem(rule["to"]))
            parts = [f"{u}:{f}" for u, f in rule["factors"].items()]
            self.table.setItem(r, 2,
                QtWidgets.QTableWidgetItem("; ".join(parts)))
        v.addWidget(self.table)

        btns = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить правило")
        self.btn_add.clicked.connect(self.add_rule)
        self.btn_del = QtWidgets.QPushButton("Удалить правило")
        self.btn_del.clicked.connect(self.del_rule)
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_del)
        v.addLayout(btns)

        self.chk_disable = QtWidgets.QCheckBox("Отключить")
        self.chk_disable.setChecked(not self.cfg.get("enabled", True))
        self.chk_disable.stateChanged.connect(self._update_ui_state)
        v.addWidget(self.chk_disable)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        self._update_ui_state()
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())

    def _update_ui_state(self):
        disabled = self.chk_disable.isChecked()
        self.table.setDisabled(disabled)
        self.btn_add.setDisabled(disabled)
        self.btn_del.setDisabled(disabled)

    def accept(self):
        self.cfg["enabled"] = not self.chk_disable.isChecked()

        rules = []
        for i in range(self.table.rowCount()):
            col = self.table.item(i, 0).text().strip()
            to  = self.table.item(i, 1).text().strip()
            raw = self.table.item(i, 2).text().strip()
            fct = {}
            for part in raw.split(";"):
                if ":" in part:
                    u, val = part.split(":")
                    try:
                        fct[u.strip().lower()] = float(val)
                    except:
                        pass
            if col and to and fct:
                rules.append({"column": col, "to": to, "factors": fct})
        self.cfg["rules"] = rules

        super().accept()


class SheetsDialog(QtWidgets.QDialog):
    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.setWindowTitle("Листы")
        self.resize(350, 450)
        self.cfg = cfg
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)
        self.table = QtWidgets.QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels([
            "Целевое", "Синонимы (через запятую)", "Не объединять"
        ])
        for rule in self.cfg.get("rules", []):
            i = self.table.rowCount(); self.table.insertRow(i)
            self.table.setItem(i,0, QtWidgets.QTableWidgetItem(rule["target"]))
            self.table.setItem(i,1, QtWidgets.QTableWidgetItem(",".join(rule["synonyms"])))
            chk = QtWidgets.QTableWidgetItem()
            chk.setCheckState(
                QtCore.Qt.Checked if rule.get("no_merge", False)
                                 else QtCore.Qt.Unchecked
            )
            self.table.setItem(i,2, chk)
        v.addWidget(self.table)

        btns = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить правило")
        self.btn_add.clicked.connect(self.add_rule)
        self.btn_del = QtWidgets.QPushButton("Удалить правило")
        self.btn_del.clicked.connect(self.del_rule)
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_del)
        v.addLayout(btns)

        self.chk_disable = QtWidgets.QCheckBox("Отключить")
        self.chk_disable.setChecked(not self.cfg.get("enabled", True))
        v.addWidget(self.chk_disable)
        self.chk_disable.stateChanged.connect(self._update_ui_state)

        self.grp_fuzzy = QtWidgets.QGroupBox("Настройки FuzzyWuzzy")
        fl = QtWidgets.QFormLayout(self.grp_fuzzy)

        self.spin_thr = QtWidgets.QSpinBox()
        self.spin_thr.setRange(0, 100)
        self.spin_thr.setValue(self.cfg.get("threshold", 90))
        fl.addRow("Порог объединения (0–100):", self.spin_thr)

        self.chk_auto_merge = QtWidgets.QCheckBox("Объединять автоматически")
        self.chk_auto_merge.setChecked(self.cfg.get("auto_merge", True))
        fl.addRow(self.chk_auto_merge)

        v.addWidget(self.grp_fuzzy)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        self._update_ui_state()
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())
    
    def _update_ui_state(self):
        disabled = self.chk_disable.isChecked()
        self.table.setDisabled(disabled)
        self.btn_add.setDisabled(disabled)
        self.btn_del.setDisabled(disabled)
        self.grp_fuzzy.setDisabled(disabled)

    def accept(self):
        self.cfg["enabled"] = not self.chk_disable.isChecked()
        self.cfg["threshold"] = self.spin_thr.value()
        self.cfg["auto_merge"] = self.chk_auto_merge.isChecked()
        rules = []
        for i in range(self.table.rowCount()):
            tgt = self.table.item(i,0).text().strip()
            syn = [s.strip() for s in self.table.item(i,1).text().split(",") if s.strip()]
            no  = self.table.item(i,2).checkState()==QtCore.Qt.Checked
            rules.append({"target":tgt,"synonyms":syn,"no_merge":no})
        self.cfg["rules"] = rules
        super().accept()

class WordFilterDialog(QtWidgets.QDialog):
    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.setWindowTitle("Фильтр слов")
        self.cfg = cfg
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)
        self.table = QtWidgets.QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Слово","Удалять всю строку"])
        for rule in self.cfg.get("rules", []):
            r = self.table.rowCount(); self.table.insertRow(r)
            self.table.setItem(r,0, QtWidgets.QTableWidgetItem(rule["word"]))
            chk = QtWidgets.QTableWidgetItem()
            chk.setFlags(chk.flags() | QtCore.Qt.ItemIsUserCheckable)
            chk.setCheckState(
                QtCore.Qt.Checked if rule.get("delete_row", False)
                                 else QtCore.Qt.Unchecked
            )
            self.table.setItem(r,1, chk)
        v.addWidget(self.table)

        btns = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить правило")
        self.btn_add.clicked.connect(self.add_rule)
        self.btn_del = QtWidgets.QPushButton("Удалить правило")
        self.btn_del.clicked.connect(self.del_rule)
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_del)
        v.addLayout(btns)

        self.chk_disable = QtWidgets.QCheckBox("Отключить")
        self.chk_disable.setChecked(not self.cfg.get("enabled", True))
        v.addWidget(self.chk_disable)
        self.chk_disable.stateChanged.connect(self._update_ui_state)


        self.grp_fuzzy = QtWidgets.QGroupBox("Настройки FuzzyWuzzy")
        fl = QtWidgets.QFormLayout(self.grp_fuzzy)
        self.spin_thr = QtWidgets.QSpinBox()
        self.spin_thr.setRange(0, 100)
        self.spin_thr.setValue(self.cfg.get("threshold", 90))
        fl.addRow("Порог совпадения (0–100):", self.spin_thr)
        self.chk_auto = QtWidgets.QCheckBox("Удалять автоматически")
        self.chk_auto.setChecked(self.cfg.get("auto_merge", True))
        fl.addRow(self.chk_auto)

        v.addWidget(self.grp_fuzzy)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        self._update_ui_state()
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())

    def _update_ui_state(self):
        disabled = self.chk_disable.isChecked()
        self.table.setDisabled(disabled)
        self.btn_add.setDisabled(disabled)
        self.btn_del.setDisabled(disabled)
        self.grp_fuzzy.setDisabled(disabled)

    def accept(self):
        self.cfg["enabled"] = not self.chk_disable.isChecked()
        self.cfg["threshold"]  = self.spin_thr.value()
        self.cfg["auto_merge"] = self.chk_auto.isChecked()
        rules = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i,0)
            if not item or not item.text().strip():
                continue
            w = item.text().strip()
            delete = self.table.item(i,1).checkState()==QtCore.Qt.Checked
            rules.append({"word": w, "delete_row": delete})
        self.cfg["rules"] = rules
        super().accept()


class ColumnWordFilterDialog(QtWidgets.QDialog):
    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.setWindowTitle("Фильтр слов по столбцам")
        self.cfg = cfg
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)
        self.table = QtWidgets.QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Слово","Удалять всю строку"])
        for rule in self.cfg.get("rules", []):
            r = self.table.rowCount(); self.table.insertRow(r)
            self.table.setItem(r,0, QtWidgets.QTableWidgetItem(rule["word"]))
            chk = QtWidgets.QTableWidgetItem()
            chk.setFlags(chk.flags() | QtCore.Qt.ItemIsUserCheckable)
            chk.setCheckState(
                QtCore.Qt.Checked if rule.get("delete_row", False)
                                 else QtCore.Qt.Unchecked
            )
            self.table.setItem(r,1, chk)
        v.addWidget(self.table)

        btns = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить правило")
        self.btn_add.clicked.connect(self.add_rule)
        self.btn_del = QtWidgets.QPushButton("Удалить правило")
        self.btn_del.clicked.connect(self.del_rule)
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_del)
        v.addLayout(btns)

        self.chk_disable = QtWidgets.QCheckBox("Отключить")
        self.chk_disable.setChecked(not self.cfg.get("enabled", True))
        v.addWidget(self.chk_disable)
        self.chk_disable.stateChanged.connect(self._update_ui_state)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        self._update_ui_state()
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())

    def _update_ui_state(self):
        disabled = self.chk_disable.isChecked()
        self.table.setDisabled(disabled)
        self.btn_add.setDisabled(disabled)
        self.btn_del.setDisabled(disabled)

    def accept(self):
        self.cfg["enabled"] = not self.chk_disable.isChecked()
        rules = []
        for i in range(self.table.rowCount()):
            w   = self.table.item(i,0).text().strip()
            dr  = self.table.item(i,1).checkState()==QtCore.Qt.Checked
            if w:
                rules.append({"word": w, "delete_row": dr})
        self.cfg["rules"] = rules
        super().accept()


class ReplaceWordsDialog(QtWidgets.QDialog):
    def __init__(self, parent, cfg):
        super().__init__(parent)
        self.setWindowTitle("Замена слов")
        self.cfg = cfg
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(
            ["Целевое","Синонимы (через запятую)"]
        )
        for rule in self.cfg.get("rules", []):
            r = self.table.rowCount(); self.table.insertRow(r)
            self.table.setItem(r,0,
                QtWidgets.QTableWidgetItem(rule["target"]))
            self.table.setItem(r,1,
                QtWidgets.QTableWidgetItem(",".join(rule["synonyms"])))
        v.addWidget(self.table)

        h = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Добавить правило")
        self.btn_add.clicked.connect(self.add_rule)
        self.btn_del = QtWidgets.QPushButton("Удалить правило")
        self.btn_del.clicked.connect(self.del_rule)
        h.addWidget(self.btn_add)
        h.addWidget(self.btn_del)
        v.addLayout(h)

        self.chk_disable = QtWidgets.QCheckBox("Отключить")
        self.chk_disable.setChecked(not self.cfg.get("enabled", True))
        v.addWidget(self.chk_disable)
        self.chk_disable.stateChanged.connect(self._update_ui_state)

        self.grp_fuzzy = QtWidgets.QGroupBox("Настройки FuzzyWuzzy")
        fl = QtWidgets.QFormLayout(self.grp_fuzzy)
        self.spin_thr = QtWidgets.QSpinBox()
        self.spin_thr.setRange(0, 100)
        self.spin_thr.setValue(self.cfg.get("threshold", 80))
        fl.addRow("Порог объединения (0–100):", self.spin_thr)
        self.chk_auto = QtWidgets.QCheckBox("Заменять автоматически")
        self.chk_auto.setChecked(self.cfg.get("auto_replace", True))
        fl.addRow(self.chk_auto)
        v.addWidget(self.grp_fuzzy)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        self._update_ui_state()
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())

    def _update_ui_state(self):
        disabled = self.chk_disable.isChecked()
        self.table.setDisabled(disabled)
        self.btn_add.setDisabled(disabled)
        self.btn_del.setDisabled(disabled)
        self.grp_fuzzy.setDisabled(disabled)

    def accept(self):
        self.cfg["enabled"] = not self.chk_disable.isChecked()
        rules = []
        for i in range(self.table.rowCount()):
            tgt_item = self.table.item(i,0)
            syn_item = self.table.item(i,1)
            if not tgt_item or not syn_item:
                continue
            tgt = tgt_item.text().strip()
            syns= [s.strip() for s in syn_item.text().split(",") if s.strip()]
            if tgt and syns:
                rules.append({"target": tgt, "synonyms": syns})
        self.cfg["rules"]       = rules
        self.cfg["threshold"]   = self.spin_thr.value()
        self.cfg["auto_replace"]= self.chk_auto.isChecked()
        self.cfg["threshold"]     = self.spin_thr.value()
        self.cfg["auto_replace"]  = self.chk_auto.isChecked()
        super().accept()


class RulesManagerDialog(QtWidgets.QDialog):
    def __init__(self, parent, rules):
        super().__init__(parent)
        self.setWindowTitle("Управление правилами")
        self.rules = rules
        self._init_ui()

    def _init_ui(self):
        v = QtWidgets.QVBoxLayout(self)
        btns = [
            ("Столбцы",               self.open_columns),
            ("Единицы измерения",     self.open_units),
            ("Листы",                 self.open_sheets),
            ("Фильтр слов",           self.open_filter),
            ("Замена слов",           self.open_replace),
            ("Фильтр слов по столбцам", self.open_column_filter)
        ]
        for txt, fn in btns:
            b = QtWidgets.QPushButton(txt)
            b.clicked.connect(fn)
            v.addWidget(b)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save |
            QtWidgets.QDialogButtonBox.Cancel
        )
        localize_buttonbox(bb)
        bb.accepted.connect(self.save_and_close)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
    
    def add_rule(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col in range(self.table.columnCount()):
            self.table.setItem(row, col, QtWidgets.QTableWidgetItem(""))

    def del_rule(self):
        selected = self.table.selectionModel().selectedRows()
        for idx in sorted(selected, key=lambda x: x.row(), reverse=True):
            self.table.removeRow(idx.row())

    def open_columns(self):
        dlg = ColumnsDialog(self, self.rules["column_rules"])
        dlg.exec()

    def open_units(self):
        dlg = UnitsDialog(self, self.rules["unit_rules"])
        dlg.exec()

    def open_sheets(self):
        dlg = SheetsDialog(self, self.rules["sheet_rules"])
        dlg.exec()

    def open_filter(self):
        dlg = WordFilterDialog(self, self.rules["word_filter"])
        dlg.exec()

    def open_replace(self):
        dlg = ReplaceWordsDialog(self, self.rules["word_replace"])
        dlg.exec()

    def open_column_filter(self):
        dlg = ColumnWordFilterDialog(self, self.rules["column_word_filter"])
        dlg.exec()

    def save_and_close(self):
        save_rules(self.rules)
        self.accept()
