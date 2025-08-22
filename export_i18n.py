import os
import polib
from openpyxl import Workbook
from collections import OrderedDict


ZH_CN = "zh_CN"


def get_pot_msgids():
    """从templates目录提取所有pot文件的唯一msgid（保持顺序）"""
    msgids = OrderedDict()
    pot_dir = "./templates"

    for filename in os.listdir(pot_dir):
        if filename.endswith(".pot"):
            po = polib.pofile(os.path.join(pot_dir, filename))
            for entry in po:
                if entry.msgid and entry.msgid not in msgids:
                    msgids[entry.msgid] = True
    return list(msgids.keys())


def get_po_translations():
    """收集各语言翻译内容"""
    translations = {}

    for lang in os.listdir("."):
        lang_dir = os.path.join(".", lang)
        if os.path.isdir(lang_dir) and lang != "templates" and lang != ZH_CN:
            po_path = os.path.join(lang_dir, f"{lang}.po")
            if os.path.exists(po_path):
                po = polib.pofile(po_path)
                translations[lang] = {entry.msgid: entry.msgstr for entry in po}
    return translations


def generate_spreadsheet():
    """生成Excel文件"""
    msgids = get_pot_msgids()
    translations = get_po_translations()
    languages = list(translations.keys())

    wb = Workbook()
    ws = wb.active

    # 写入标题行
    headers = [ZH_CN] + languages
    ws.append(headers)

    # 写入翻译内容
    for msgid in msgids:
        row = [msgid]
        for lang in languages:
            row.append(translations[lang].get(msgid, ""))
        ws.append(row)

    wb.save("i18n.xlsx")


if __name__ == "__main__":
    generate_spreadsheet()
