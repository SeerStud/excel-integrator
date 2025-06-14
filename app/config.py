import os
import json

RULES_FILE = os.path.join(os.path.dirname(__file__), '..', 'rules.json')

def load_rules():
    if os.path.exists(RULES_FILE):
        with open(RULES_FILE, 'r', encoding='utf-8') as f:
            rules = json.load(f)

        for section in (
            "column_rules","unit_rules","sheet_rules",
            "word_filter","word_replace","column_word_filter"
        ):
            sec = rules.get(section, {})
            sec.setdefault("enabled", True)
            rules[section] = sec

        rules.setdefault("unit_rules", {})
        rules["unit_rules"].setdefault("no_unit_to_header", False)

        return rules

    return {
        "column_rules":       {"rules":[],"threshold":80,"auto_merge":True,"enabled":True},
        "unit_rules":         {"rules":[],"threshold":80,"auto_merge":True,"enabled":True,"no_unit_to_header":False},
        "sheet_rules":        {"rules":[],"threshold":90,"auto_merge":True,"enabled":True},
        "word_filter":        {"rules":[],"threshold":60,"auto_merge":True,"enabled":True},
        "word_replace":       {"rules":[],"threshold":80,"auto_replace":True,"enabled":True},
        "column_word_filter": {"rules":[],"enabled":True},
        "skip_rows_keywords": []
    }

def save_rules(rules):
    with open(RULES_FILE, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=4)