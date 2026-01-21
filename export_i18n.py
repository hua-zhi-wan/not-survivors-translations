import os
import polib
from openpyxl import Workbook
from collections import OrderedDict


ZH_CN = "zh_CN"


def get_pot_entries():
    """从templates目录提取所有pot文件的唯一条目（保持顺序），同时考虑msgid和msgctxt"""
    entries = OrderedDict()
    pot_dir = "./templates"

    for filename in os.listdir(pot_dir):
        if filename.endswith(".pot"):
            po = polib.pofile(os.path.join(pot_dir, filename))
            for entry in po:
                if entry.msgid:
                    # 使用 (msgctxt, msgid) 作为键，确保不同上下文的相同msgid被视为不同条目
                    key = (entry.msgctxt, entry.msgid)
                    if key not in entries:
                        entries[key] = entry
    return list(entries.values())


def get_po_translations():
    """收集各语言翻译内容，考虑msgctxt"""
    translations = {}

    for lang in os.listdir("."):
        lang_dir = os.path.join(".", lang)
        if os.path.isdir(lang_dir) and lang != "templates" and lang != ZH_CN:
            po_path = os.path.join(lang_dir, f"{lang}.po")
            if os.path.exists(po_path):
                po = polib.pofile(po_path)
                # 使用 (msgctxt, msgid) 作为键，确保不同上下文的相同msgid被视为不同条目
                translations[lang] = {(entry.msgctxt, entry.msgid): entry.msgstr for entry in po}
    return translations


def generate_spreadsheet():
    """生成Excel文件"""
    pot_entries = get_pot_entries()
    translations = get_po_translations()
    languages = list(translations.keys())

    wb = Workbook()
    ws = wb.active

    # 写入标题行
    headers = [ZH_CN] + languages
    ws.append(headers)

    # 写入翻译内容
    for entry in pot_entries:
        msgid = entry.msgid
        row = [msgid]
        for lang in languages:
            # 使用 (msgctxt, msgid) 作为键来获取翻译
            key = (entry.msgctxt, msgid)
            row.append(translations[lang].get(key, ""))
        ws.append(row)

    wb.save("i18n.xlsx")


if __name__ == "__main__":
    generate_spreadsheet()
