import re
import pandas as pd
import spacy
from rapidfuzz import fuzz
from datetime import datetime
from PySide6 import QtWidgets

_nlp = spacy.load("ru_core_news_lg")

def split_src(col: str):
    if col.startswith("__src"):
        parts = col.split("__", 2)
        return parts[1], parts[2]
    return None, col

def apply_word_replace(df: pd.DataFrame, cfg: dict, log: list) -> pd.DataFrame:
    if not cfg.get("enabled", True):
        return df
    rules = cfg.get("rules", [])
    threshold = cfg.get("threshold", 80)
    auto = cfg.get("auto_replace", True)

    df2 = df.copy()

    exact_map = {r["target"].lower(): r["target"] for r in rules}

    def _replace_cell(val, idx, col):
        if pd.isna(val) or not isinstance(val, str):
            return val
        parts = re.split(r'(\W+)', val)
        changed = False
        new_parts = []

        for token in parts:
            low = token.lower()
            if low in exact_map:
                new_parts.append(token)
                continue

            for r in rules:
                if low in [s.lower() for s in r["synonyms"]]:
                    new_parts.append(r["target"])
                    log.append(f"Замена слов: '{token}' → '{r['target']}' в [{idx},'{col}']")
                    changed = True
                    break
            else:
                best_score, best_target = 0, None
                for r in rules:
                    for cand in (r["target"], *r["synonyms"]):
                        sc = fuzz.token_set_ratio(low, cand.lower())
                        if sc > best_score:
                            best_score, best_target = sc, r["target"]
                if best_score >= threshold and best_target:
                    if auto:
                        new_val = best_target
                    else:
                        dlg = QtWidgets.QDialog()
                        dlg.setWindowTitle("Замена слова")
                        form = QtWidgets.QFormLayout(dlg)
                        chk = QtWidgets.QCheckBox(f"Заменить '{token}' → '{best_target}' ({round(best_score)}%)")
                        chk.setChecked(True)
                        form.addRow(chk)
                        name_edit = QtWidgets.QLineEdit(best_target)
                        form.addRow("Новое слово:", name_edit)
                        bb = QtWidgets.QDialogButtonBox(
                            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
                        )
                        bb.accepted.connect(dlg.accept)
                        bb.rejected.connect(dlg.reject)
                        form.addRow(bb)
                        if dlg.exec() == QtWidgets.QDialog.Accepted and chk.isChecked():
                            new_val = name_edit.text().strip() or best_target
                        else:
                            new_val = token
                    new_parts.append(new_val)
                    log.append(
                        f"Замена слов (fuzzy): '{token}' → '{new_val}' в [{idx},'{col}'] ({round(best_score)}%)"
                    )
                    changed = True
                else:
                    new_parts.append(token)

        return "".join(new_parts) if changed else val

    for col in df2.select_dtypes(include=["object", "string"]):
        df2[col] = df2[col].reset_index().apply(
            lambda rec: _replace_cell(rec[col], rec.name, col), axis=1
        )

    return df2

def apply_word_filter(df: pd.DataFrame, cfg: dict, log: list) -> pd.DataFrame:
    if not cfg.get("enabled", True):
        return df
    rules = cfg.get("rules", [])
    threshold = cfg.get("threshold", 60)
    df2 = df.copy()
    to_drop = set()

    for rule in rules:
        bad = rule["word"].strip().lower()
        delete_row = rule.get("delete_row", False)
        word_re = re.compile(rf"\b{re.escape(bad)}\b", re.IGNORECASE)

        for col in df2.select_dtypes(include=["object", "string"]):
            def _filter_cell(rec):
                idx = rec.name
                val = rec[col]
                if not isinstance(val, str) or not val.strip():
                    return val

                if word_re.search(val):
                    if delete_row:
                        to_drop.add(idx)
                        log.append(f'Фильтр слов: удалена строка {idx}')
                        return val
                    else:
                        log.append(f'Фильтр слов: очищена ячейка [{idx}, "{col}"]')
                        return ""

                score = fuzz.token_set_ratio(val.lower(), bad)
                if score >= threshold:
                    msg = f'{"Удалить строку" if delete_row else "Удалить слово"} «{bad}»?'
                    reply = QtWidgets.QMessageBox.question(
                        None, "Фильтр слов", msg,
                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                        QtWidgets.QMessageBox.Yes
                    )
                    if reply == QtWidgets.QMessageBox.Yes:
                        if delete_row:
                            to_drop.add(idx)
                            log.append(f'Фильтр слов (fuzzy): удалена строка {idx} ({round(score)}%)')
                            return val
                        else:
                            log.append(f'Фильтр слов (fuzzy): очищена ячейка [{idx}, "{col}"] ({round(score)}%)')
                            return ""

                return val

            df2[col] = df2[col].reset_index().apply(_filter_cell, axis=1)

    if to_drop:
        df2 = df2.drop(index=sorted(to_drop)).reset_index(drop=True)

    return df2

def extract_units_to_headers(df: pd.DataFrame, rules: dict, log: list) -> pd.DataFrame:
    unit_cfg = rules.get("unit_rules", {})
    allowed_units = {
        u.lower()
        for rule in unit_cfg.get("rules", [])
        for u in rule.get("factors", {}).keys()
    }

    pat = re.compile(r'''
        ^\s*
        (?:(?P<num>\d+\.?\d*)\s*(?P<unit1>[^\d\s]+)
        |(?P<unit2>[^\d\s]+)\s*(?P<num2>\d+\.?\d*))
        \s*$
    ''', re.IGNORECASE | re.VERBOSE)

    def normalize_cell(v):
        m = pat.match(str(v))
        if not m:
            return v
        unit = (m.group('unit1') or m.group('unit2')).lower()
        if unit not in allowed_units:
            return v
        num = m.group('num') or m.group('num2')
        return f"{num} {unit}"

    df2 = df.copy()
    for col in df2.columns:
        df2[col] = df2[col].map(normalize_cell)
    return df2

def apply_column_rules(df: pd.DataFrame, cfg: dict, log: list, col_source: dict) -> pd.DataFrame:
    if not cfg.get("enabled", True):
        return df
    tbl = df.copy()

    def split_src(col):
        if col.startswith("__src"):
            src, bare = col.split("__", 2)[1:]
            return src, bare
        return None, col

    skip_cols = {
        bare.lower()
        for rule in cfg["rules"] if rule.get("no_merge", False)
        for bare in ([rule["target"]] + rule.get("synonyms", []))
    }
    for rule in cfg["rules"]:
        if rule.get("no_merge", False):
            continue
        keys = {rule["target"].lower(), *map(str.lower, rule.get("synonyms", []))}
        found = [c for c in tbl.columns if c.lower() in keys]
        if len(found) > 1:
            log.append(f"Словарно объединены {found} → '{rule['target']}'")
            tbl[rule["target"]] = tbl[found].bfill(axis=1).iloc[:, 0]
            for c in found:
                if c != rule["target"]:
                    tbl.drop(columns=[c], inplace=True)

    use_content  = cfg.get("use_content", False)
    content_rows = cfg.get("content_rows", 10)
    alpha = cfg.get("header_weight", 0.6)

    cols = list(tbl.columns)
    i = 0
    while i < len(cols):
        base = cols[i]
        j = i + 1
        while j < len(cols):
            other = cols[j]
            if col_source.get(base) == col_source.get(other):
                j += 1
                continue

            if base.lower() in skip_cols or other.lower() in skip_cols:
                j += 1
                continue

            H = _nlp(base.lower()).similarity(_nlp(other.lower()))
            C = H
            if use_content:
                series1 = tbl[base].dropna().astype(str)
                series2 = tbl[other].dropna().astype(str)

                if len(series1) >= content_rows:
                    vals1 = series1.sample(content_rows, random_state=0).tolist()
                else:
                    vals1 = series1.tolist()
                if len(series2) >= content_rows:
                    vals2 = series2.sample(content_rows, random_state=0).tolist()
                else:
                    vals2 = series2.tolist()

                vals1 = [v for v in vals1 if re.search(r"[A-Za-zА-Яа-я]", v)]
                vals2 = [v for v in vals2 if re.search(r"[A-Za-zА-Яа-я]", v)]

                ents1 = set()
                ents2 = set()
                for v in vals1:
                    for tok in _nlp(v):
                        if tok.pos_ == "PROPN":
                            ents1.add(tok.text)
                for v in vals2:
                    for tok in _nlp(v):
                        if tok.pos_ == "PROPN":
                            ents2.add(tok.text)


                if ents1 and ents2:
                    inter = ents1 & ents2
                    union = ents1 | ents2
                    C = len(inter) / len(union)
                else:
                    C = H
            alpha = cfg.get("header_weight", 0.6)
            score = int((alpha * H + (1 - alpha) * C) * 100)

            thr = cfg.get("threshold", 80)
            hdr_ok  = (H * 100) >= thr   
            cnt_ok  = (C * 100) >= thr   
            comb_ok = score       >= thr   

            if hdr_ok or cnt_ok or comb_ok:
                if cfg.get("auto_merge", False):
                    do_merge = True
                    name = base
                else:
                    dlg = QtWidgets.QDialog()
                    dlg.setWindowTitle("Объединить столбцы?")
                    form = QtWidgets.QFormLayout(dlg)

                    chk = QtWidgets.QCheckBox(f"Объединить '{base}' + '{other}' ({score}%)")
                    chk.setChecked(True)
                    form.addRow(chk)

                    name_edit = QtWidgets.QLineEdit(base)
                    form.addRow("Имя результирующего столбца:", name_edit)

                    bb = QtWidgets.QDialogButtonBox(
                        QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
                    )
                    bb.accepted.connect(dlg.accept)
                    bb.rejected.connect(dlg.reject)
                    form.addRow(bb)

                    if dlg.exec() == QtWidgets.QDialog.Accepted and chk.isChecked():
                        do_merge = True
                        name = name_edit.text().strip() or base
                    else:
                        do_merge = False

                if do_merge:
                    log.append(f"NER объединены '{base}' + '{other}' → '{name}' ({score}%)")
                    tbl[name] = tbl[[base, other]].bfill(axis=1).iloc[:, 0]
                    for c in (base, other):
                        if c != name:
                            tbl.drop(columns=[c], inplace=True)
                    cols[i] = name
                    cols.pop(j)
                    continue

            j += 1
        i += 1

    return tbl

def apply_unit_conversions(df: pd.DataFrame, cfg: dict, log: list) -> pd.DataFrame:
    pat = re.compile(r"^\s*([\d\.]+)\s*([^\d\.\s]+)\s*$", re.IGNORECASE)

    factors = {
        unit.strip().lower(): (coef, rule["to"])
        for rule in cfg.get("rules", [])
        for unit, coef in rule.get("factors", {}).items()
    }

    def to_num(v):
        s = str(v).strip()
        m = pat.match(s)
        if m:
            num, unit = m.group(1), m.group(2).lower()
            if unit in factors:
                coef, _ = factors[unit]
                try:
                    return float(num.replace(",", ".")) * coef
                except ValueError:
                    pass
        try:
            return float(s.replace(",", "."))
        except ValueError:
            return v

    def fmt(v, unit):
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return f"{v} {unit}"

    df2 = df.copy()

    for col in df.columns:
        raw = df[col].dropna().astype(str)
        found = {
            pat.match(x.strip()).group(2).lower().rstrip('.:')
            for x in raw if pat.match(x.strip())
        }
        valid = {u for u in found if u in factors}
        if not valid:
            continue

        df2[col] = df[col].map(to_num)

        target_units = {factors[u][1] for u in valid}
        if len(target_units) == 1:
            tgt = next(iter(target_units))
            df2[col] = df2[col].map(lambda v: fmt(v, tgt) if pd.notna(v) else v)
            log.append(
                f"Единицы в ячейках столбца '{col}': форматирование значений с единицей '{tgt}'"
            )

    return df2