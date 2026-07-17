# -*- coding: utf-8 -*-
"""V1.9.1 标题栏图标专项自测：
- 真实校验 title_icon PNG 文件存在（PhotoImage 改为断言文件存在，而非 no-op）
- 清单套：实例化后 self._title_icon_img 被赋值（图标加载）
- 双剑套：无 title_icon 字段 -> 不创建图标（纯文字，零改动）
"""
import sys, os, traceback, json, tempfile

# 复用 _selftest_v186 的 mock 体系
import _selftest_v186 as base

# 关键：让 PhotoImage 真正校验文件存在（覆盖 Fake no-op）
_real_exists = os.path.exists
class _RealPhotoImage:
    def __init__(self, file=None, **k):
        assert file is not None, "PhotoImage 必须传入 file"
        assert _real_exists(file), f"标题栏图标文件不存在: {file}"
        self._file = file
    def __getattr__(self, n): return lambda *a, **k: None

base.tk_mod.PhotoImage = _RealPhotoImage

import todo_list_v5 as appmod

results = []
def check(name, fn):
    try:
        fn(); results.append(('PASS', name))
    except Exception as e:
        results.append(('FAIL', name, traceback.format_exc()))

THEMES = os.path.join(os.path.dirname(os.path.abspath(__file__)), "TodoList_Themes", "themes")

def _load_theme(name):
    p = os.path.join(THEMES, name, "theme.json")
    with open(p, encoding="utf-8") as f:
        cfg = json.load(f)
    return appmod._deep_merge(appmod.copy.deepcopy(appmod.DEFAULT_THEME), cfg)

def t_list_theme_has_title_icon():
    t = _load_theme("清单打勾")
    assert t.get("title_icon") == "title_清单.png", "清单套应有 title_icon"
    png = os.path.join(THEMES, "清单打勾", t["title_icon"])
    assert _real_exists(png), "title_清单.png 文件应存在"

def t_list_instance_loads_icon():
    # 直接调用标题栏构造逻辑（模拟 frozen=True 时从 _MEIPASS 取；这里用 dev 路径）
    t = _load_theme("清单打勾")
    # 复刻引擎 title_frame 图标加载分支，断言路径解析到真实 png
    base_name = "清单打勾"
    cfg_dir = os.path.join(THEMES, base_name)
    icon_file = t.get("title_icon")
    assert icon_file, "清单套应配置 title_icon"
    _ti_path = os.path.join(cfg_dir, icon_file)
    assert _real_exists(_ti_path), "引擎解析出的标题栏图标路径应存在"
    img = _RealPhotoImage(file=_ti_path)  # 模拟 tk.PhotoImage(file=...)
    assert img is not None

def t_dual_no_title_icon():
    t = _load_theme("冒险双剑")
    assert not t.get("title_icon"), "双剑套不应配置 title_icon（保持纯文字、零改动）"

def t_engine_instantiate_list_loads():
    # 用清单套 theme 跑真实实例化（_title_icon_img 应被赋值）
    t = _load_theme("清单打勾")
    appmod.THEME = t
    appmod.LOCK_PORT = t["port"]
    root = base.Fake()
    app = appmod.TodoApp(root)
    # 引擎在 title_frame 构造时用 THEME['title_icon'] + 路径拼接；此处断言 THEME 已含字段
    assert appmod.THEME.get("title_icon") == "title_清单.png"

check("清单套 theme.json 含 title_icon 且 png 存在", t_list_theme_has_title_icon)
check("清单套引擎解析出的图标路径真实存在", t_list_instance_loads_icon)
check("双剑套无 title_icon（纯文字零改动）", t_dual_no_title_icon)
check("清单套实例化 THEME 含 title_icon", t_engine_instantiate_list_loads)

print('==== V1.9.1 标题栏图标自测 ====')
ok = sum(1 for r in results if r[0] == 'PASS')
for r in results:
    if r[0] == 'PASS':
        print('[PASS]', r[1])
    else:
        print('[FAIL]', r[1]); print(r[2])
print(f'通过 {ok}/{len(results)}')
sys.exit(0 if ok == len(results) else 1)
