import os
import polib
from openpyxl import load_workbook


def update_translations():
    """更新各语言PO文件"""
    wb = load_workbook("i18n.xlsx")
    ws = wb.active

    # 读取语言列表
    languages = [cell.value for cell in ws[1]][0:]  # 跳过zh_CN列

    # 构建翻译数据 {msgid: {lang: translation}}
    translations = {}
    for row in ws.iter_rows(min_row=2):
        msgid = row[0].value
        if not msgid:
            continue

        translations[msgid] = {
            lang: row[i].value or "" for i, lang in enumerate(languages)
        }

    # 更新各语言文件
    for lang in languages:
        lang_dir = os.path.join(".", lang)
        po_path = os.path.join(lang_dir, f"{lang}.po")

        if not os.path.exists(po_path):
            print(f"警告：{po_path} 不存在，已跳过")
            continue

        po = polib.pofile(po_path)
        existing = {entry.msgid: entry for entry in po}

        # 更新现有条目并添加新条目
        for msgid, trans in translations.items():
            msgstr = trans.get(lang, "")

            if msgid in existing:
                existing[msgid].msgstr = msgstr
            else:
                new_entry = polib.POEntry(msgid=msgid, msgstr=msgstr)
                po.append(new_entry)

        po.save(po_path)


if __name__ == "__main__":
    update_translations()
