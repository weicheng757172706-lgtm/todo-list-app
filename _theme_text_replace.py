# -*- coding: utf-8 -*-
"""把源码里硬编码的界面文案精确替换为 self.T['KEY'] 主题引用。
双剑套用 DEFAULT_THEME.text（冒险文案），清单套在 theme.json 覆盖。
所有 old 均为带引号的完整字面量，避免 emoji 串误伤 form label。
"""
import io, sys

p = "todo_list_v5.py"
s = open(p, encoding="utf-8").read()

REPL = [
    # ── 列名（单引号，无冒号；同时作 Treeview 列 key）──
    ("'📝 冒险描述'", "self.T['col_desc']"),
    ("'🗺️ 冒险类型'", "self.T['col_type']"),
    ("'🗺️冒险类型'", "self.T['col_type']"),   # 无空格变体（completed tree）
    ("'⚡ 紧急等级'", "self.T['col_priority']"),
    ("'⚔️ 冒险者'", "self.T['col_owner']"),
    ("'🏷️ 标签'", "self.T['col_tag']"),
    ("'📊 进度'", "self.T['col_progress']"),
    ("'⏰ 领取时间'", "self.T['col_receive']"),
    ("'🏁 通关时间'", "self.T['col_complete_time']"),
    ("'⏱️ 冒险时长'", "self.T['col_duration']"),
    ("'🚀 进行中'", "self.T['stat_ongoing']"),
    ("'🏁 已通关'", "self.T['stat_done']"),
    ("'📊 合计'", "self.T['stat_total']"),
    # ── 表单标签（双引号，带冒号）──
    ('"🎯 冒险描述:"', "self.T['lbl_desc_add']"),
    ('"📝 冒险描述:"', "self.T['lbl_desc_edit']"),
    ('"⚡ 紧急等级:"', "self.T['lbl_priority']"),
    ('"🗺️ 冒险类型:"', "self.T['lbl_type']"),
    ('"⚔️ 冒险者:"', "self.T['lbl_owner']"),
    # ── 顶部 tab ──
    ('"🚀 进行中的冒险"', "self.T['tab_pending']"),
    ('"🏁 已通关的冒险"', "self.T['tab_completed']"),
    ('"📊 冒险战绩"', "self.T['tab_stats']"),
    # ── 工具栏按钮 ──
    ('"🎯 标记通关"', "self.T['btn_mark']"),
    ('"🗑 放弃冒险"', "self.T['btn_abandon']"),
    ('"✏ 编辑冒险"', "self.T['btn_edit']"),
    ('"🔄 刷新冒险"', "self.T['btn_refresh']"),
    ('"🗑 清空冒险"', "self.T['btn_clear']"),
    ('"📤 导出战绩"', "self.T['btn_export']"),
    # ── 统计区 frame ──
    ('"⚡ 按紧急等级"', "self.T['stat_priority']"),
    ('"🗺️ 按冒险类型"', "self.T['stat_category']"),
    # ── 统计 tree 列（整行，保险）──
    ("('⚡ 紧急等级', '🚀 进行中', '🏁 已通关', '📊 合计')",
     "(self.T['col_priority'], self.T['stat_ongoing'], self.T['stat_done'], self.T['stat_total'])"),
    ("('🗺️ 冒险类型', '🚀 进行中', '🏁 已通关', '📊 合计')",
     "(self.T['col_type'], self.T['stat_ongoing'], self.T['stat_done'], self.T['stat_total'])"),
    ("{'🗺️ 冒险类型': 135, '🚀 进行中': 95, '🏁 已通关': 95, '📊 合计': 95}",
     "{self.T['col_type']: 135, self.T['stat_ongoing']: 95, self.T['stat_done']: 95, self.T['stat_total']: 95}"),
    # ── 排序下拉 ──
    ('values=["默认顺序", "按紧急程度", "按创建时间", "按任务描述", "按通关时间"]',
     "values=[self.T['sort_default'], self.T['sort_priority'], self.T['sort_created'], self.T['sort_desc'], self.T['sort_completed']]"),
    ('if criteria == "默认顺序":', "if criteria == self.T['sort_default']:"),
    ('elif criteria == "按紧急程度":', "elif criteria == self.T['sort_priority']:"),
    # ── 导出 ──
    ('default_filename = f"{selected_week} 已通关冒险列表.xlsx"',
     "default_filename = self.T['export_file_week'].format(week=selected_week)"),
    ('default_filename = f"{iso_year}年第{iso_week}周 已通关冒险列表.xlsx"',
     "default_filename = self.T['export_file_iso'].format(year=iso_year, week=iso_week)"),
    ('title="📤 导出战绩"', "title=self.T['export_title']"),
    ('ws.title = "🏁 已通关冒险"', "ws.title = self.T['export_sheet']"),
    ('headers = ["🔢 编号", "📝 冒险描述", "🗺️ 冒险类型", "⚡ 紧急等级", "⚔️ 冒险者", "🏷️ 标签", "🏁 通关时间", "📅 领取时间", "⏱️ 冒险时长"]',
     "headers = [self.T['col_id_emoji'], self.T['col_desc'], self.T['col_type'], self.T['col_priority'], self.T['col_owner'], self.T['col_tag'], self.T['col_complete_time'], self.T['col_receive'], self.T['col_duration']]"),
    ('if col_idx == 2:  # 📝 冒险描述 - left align', "if col_idx == 2:  # 描述列左对齐"),
    # ── 状态栏 / 校验 / 弹窗 ──
    ('self.set_status("🎯 请输入冒险描述...")', "self.set_status(self.T['status_enter_desc'])"),
    ('messagebox.showwarning("警告", "🎯 请输入冒险描述！")',
     'messagebox.showwarning("警告", self.T[\'warn_enter_desc\'])'),
    ('"冒险者名称不能超过100个字符！"', "self.T['valid_owner']"),
    # ── 编辑对话框标题 ──
    ('dialog.title(f"✏️ 编辑冒险 #{task_id}")',
     'dialog.title(f"{self.T[\'edit_dlg_title\']}{task_id}")'),
    # ── auto_resize 子串判断改语义 ──
    ("if '冒险类型' in col:", "if col == self.T['col_type']:"),
    ("max_allowed_width = 500 if '描述' in header_text or 'description' in col.lower() else 300",
     "max_allowed_width = 500 if col == self.T['col_desc'] or 'description' in col.lower() else 300"),
    # ── 成长硬编码级别名 ──
    ('exp_card, text="菜鸟路人 | Lv.1",', 'exp_card, text=f"{self.get_wisdom_title(1)} | Lv.1",'),
    ('exp_card, text="0/60 智慧 → 见习旅者 (Lv.2)",',
     'exp_card, text=f"0/{self.get_wisdom_for_level(2)[1]} 智慧 → {self.get_wisdom_title(2)} (Lv.2)",'),
]

miss = []
for old, new in REPL:
    c = s.count(old)
    if c == 0:
        miss.append(old)
    s = s.replace(old, new)

open(p, "w", encoding="utf-8").write(s)
print(f"替换完成：{len(REPL)} 条规则，未命中 {len(miss)} 条")
for m in miss:
    print("  未命中:", repr(m))
