# -*- coding: utf-8 -*-
"""仅修复 DEFAULT_THEME.text 字典块内被 _theme_text_replace 误伤的 self.T[...] 值。
只处理 "text": { 到 "levels": [ 之间的区域，不动源码界面代码的 self.T 引用。
"""
p = "todo_list_v5.py"
lines = open(p, encoding="utf-8").read().split("\n")
start = end = None
for i, l in enumerate(lines):
    if l.strip() == '"text": {':
        start = i
    elif start is not None and l.strip() == '"levels": [':
        end = i
        break
assert start is not None and end is not None, "未定位 text 块"

MAP = {
    'tab_pending': '🚀 进行中的冒险', 'tab_completed': '🏁 已通关的冒险', 'tab_stats': '📊 冒险战绩',
    'btn_mark': '🎯 标记通关', 'btn_abandon': '🗑 放弃冒险', 'btn_edit': '✏ 编辑冒险',
    'btn_refresh': '🔄 刷新冒险', 'btn_clear': '🗑 清空冒险', 'btn_export': '📤 导出战绩',
    'lbl_desc_add': '🎯 冒险描述:', 'lbl_desc_edit': '📝 冒险描述:', 'lbl_type': '🗺️ 冒险类型:',
    'lbl_priority': '⚡ 紧急等级:', 'lbl_owner': '⚔️ 冒险者:',
    'stat_priority': '⚡ 按紧急等级', 'stat_category': '🗺️ 按冒险类型',
    'export_title': '📤 导出战绩', 'valid_owner': '冒险者名称不能超过100个字符！',
}
fixed = 0
for i in range(start, end):
    for k, v in MAP.items():
        pat = f"self.T['{k}']"
        if pat in lines[i]:
            lines[i] = lines[i].replace(pat, f'"{v}"')
            fixed += 1
open(p, "w", encoding="utf-8").write("\n".join(lines))
print(f"text 块内修复 {fixed} 处（区域行 {start+1}..{end}）")
