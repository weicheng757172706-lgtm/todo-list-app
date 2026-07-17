# -*- coding: utf-8 -*-
"""V1.9.2 专项自测：界面 emoji/文案主题化（清单套「清单勾选风」）。
1) 扫描引擎源码所有 self.T['key'] 引用，确认默认 THEME 均存在（防运行时 KeyError）
2) 直接调用所有 .format 模板，确认占位符数量匹配
3) 实例化双剑套 + 清单套，验证清单套解析出方案 A 工作风文案、双剑套保持冒险味
"""
import sys, os, re, json

# 复用 v186 的 mock 体系（openpyxl 子模块处理正确）
import _selftest_v186 as base
import todo_list_v5 as appmod

def check(name, cond):
    print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
    return cond

default_text = dict(appmod.THEME["text"])  # 捕获默认（双剑冒险风）

# 1) 扫描 self.T['key']
src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "TodoList_Themes", "engine", "todo_list_v5.py"), encoding="utf-8").read()
keys = sorted(set(re.findall(r"self\.T\['([^']+)'\]", src)))
missing = [k for k in keys if k not in default_text]
ok1 = check(f"self.T 键全部存在于默认 text（共 {len(keys)} 个，缺 {len(missing)}）", not missing)
if missing: print("    缺失:", missing)

# 2) .format 占位符校验
fmt_problems = []
trials = {
    "status_task_received": ("x",),
    "status_task_restored": ("a", "b"),
    "exp_daily": (1, 6),
    "confirm_abandon": ("x",),
    "confirm_del_done": ("x", "y"),
    "status_task_updated": ("x",),
}
for k, args in trials.items():
    try:
        default_text[k].format(*args)
    except Exception as e:
        fmt_problems.append((k, str(e)))
ok2 = check(f"所有 .format 模板调用成功（问题 {len(fmt_problems)}）", not fmt_problems)
if fmt_problems: print("    ", fmt_problems)

# 3) 双剑套实例化
root = base.tk_mod.Tk()
app = appmod.TodoApp(root)
ok3 = check("双剑套默认实例化成功", app is not None)

# 4) 清单套覆盖
THEMES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "TodoList_Themes", "themes")
t = appmod._deep_merge(appmod.copy.deepcopy(appmod.DEFAULT_THEME),
                       json.load(open(os.path.join(THEMES, "清单打勾", "theme.json"), encoding="utf-8")))
appmod.THEME = t
root2 = base.tk_mod.Tk()
app2 = appmod.TodoApp(root2)
T2 = app2.T
expect = {
    "col_type": "📋 任务类型", "col_owner": "👤 责任人", "col_priority": "⚠️ 紧急程度",
    "col_receive": "🕐 创建时间", "col_complete_time": "✅ 完成时间",
    "tab_pending": "📋 待完成", "tab_completed": "✅ 已完成", "btn_mark": "✅ 标记完成",
    "btn_edit": "✏️ 编辑任务", "exp_card_title": "📈 本周进展", "done_count_col": "✅ 完成数",
    "ach_first_title": "🎉 初次任务！", "ach_ten_title": "🏆 小有所成",
    "status_ready_icon": "📋 准备开始新任务！",
    "confirm_del_done": "🗑️ 确定要删除该已完成记录吗？\n\n任务: {}\n\n⚠️ 此操作将扣回 {} 智慧（历史总智慧同步减少），本周等级可能下降",
}
ok4 = True
for k, v in expect.items():
    good = T2.get(k) == v
    ok4 = ok4 and good
    check(f"清单套 text['{k}'] == {v!r}", good)
    if not good: print(f"      实际: {T2.get(k)!r}")

# 5) 双剑套保持冒险味
ok5 = check("双剑套 col_type 仍为 🗺️ 冒险类型（零改动）", default_text["col_type"] == "🗺️ 冒险类型")

all_ok = ok1 and ok2 and ok3 and ok4 and ok5
print("\n===== V1.9.2 自测:", "全部 PASS" if all_ok else "存在 FAIL", "=====")
sys.exit(0 if all_ok else 1)
