# -*- coding: utf-8 -*-
"""V1.8.6 主题化引擎无头自测：mock tkinter/openpyxl，验证 THEME 加载与核心实例化不崩。"""
import sys, types, os, tempfile, traceback

class Fake:
    def __init__(self, *a, **k): pass
    def __getattr__(self, name):
        return Fake()
    def __call__(self, *a, **k):
        return Fake()
    def __setitem__(self, k, v): pass
    def __getitem__(self, k): return Fake
    def __iter__(self): return iter([])
    def __bool__(self): return False
    def __contains__(self, k): return False
    # 常用方法（返回合理值，避免 KeyError/TypeError）
    def theme_names(self, *a, **k): return ['default']
    def theme_use(self, *a, **k): return None
    def configure(self, *a, **k): pass
    def config(self, *a, **k): pass
    def cget(self, *a, **k): return ''
    def after(self, *a, **k): return 0
    def mainloop(self, *a, **k): pass
    def update(self, *a, **k): pass
    def update_idletasks(self, *a, **k): pass
    def destroy(self, *a, **k): pass
    def geometry(self, *a, **k): pass
    def title(self, *a, **k): pass
    def iconbitmap(self, *a, **k): pass
    def protocol(self, *a, **k): pass
    def bind(self, *a, **k): pass
    def unbind(self, *a, **k): pass
    def focus_force(self, *a, **k): pass
    def lift(self, *a, **k): pass
    def grid(self, *a, **k): pass
    def pack(self, *a, **k): pass
    def place(self, *a, **k): pass
    def insert(self, *a, **k): return 'item'
    def get_children(self, *a, **k): return []
    def delete(self, *a, **k): pass
    def selection(self, *a, **k): return ()
    def item(self, *a, **k): return {'values': ['', '', '', '', '', '', '', '', '', '']}
    def set(self, *a, **k): pass
    def heading(self, *a, **k): pass
    def column(self, *a, **k): pass
    def tag_configure(self, *a, **k): pass
    def add(self, *a, **k): pass
    def tab(self, *a, **k): return {}
    def select(self, *a, **k): pass
    def current(self, *a, **k): return 0
    def winfo_width(self, *a, **k): return 100
    def winfo_height(self, *a, **k): return 100
    def winfo_x(self, *a, **k): return 0
    def winfo_y(self, *a, **k): return 0
    def coords(self, *a, **k): return [0, 0]
    def create_rectangle(self, *a, **k): return 1
    def create_line(self, *a, **k): return 1
    def create_oval(self, *a, **k): return 1
    def create_text(self, *a, **k): return 1
    def itemconfig(self, *a, **k): pass
    def add_command(self, *a, **k): pass
    def add_cascade(self, *a, **k): pass
    def add_separator(self, *a, **k): pass
    def post(self, *a, **k): pass
    def clipboard_clear(self, *a, **k): pass
    def clipboard_append(self, *a, **k): pass
    def index(self, *a, **k): return 0
    def get(self, *a, **k): return ''
    def winfo_exists(self, *a, **k): return 1

class _Var:
    def __init__(self, *a, **k): self._v = (a[0] if a else '')
    def get(self): return self._v
    def set(self, v): self._v = v
    def trace(self, *a, **k): return 't'
    def __getattr__(self, n): return Fake()
class MyStrVar(_Var): pass
class MyIntVar(_Var):
    def get(self):
        try: return int(self._v)
        except: return 0
class MyDoubleVar(_Var):
    def get(self):
        try: return float(self._v)
        except: return 0.0
class MyBoolVar(_Var):
    def get(self): return bool(self._v)

class _Mod(types.ModuleType):
    """带兜底 __getattr__ 的 mock 模块：任何未显式设置的属性都返回 Fake（类）。"""
    def __getattr__(self, name):
        return Fake

def make_mod():
    return _Mod('m')

# tkinter
tk_mod = make_mod()
for cls in ['Tk','Toplevel','Widget','Frame','Label','Button','Entry','LabelFrame',
            'Checkbutton','Canvas','Message','Scale','Scrollbar','Listbox','Menu',
            'Menubutton','PanedWindow','Spinbox','Text','PhotoImage']:
    setattr(tk_mod, cls, Fake)
tk_mod.StringVar = MyStrVar
tk_mod.IntVar = MyIntVar
tk_mod.DoubleVar = MyDoubleVar
tk_mod.BooleanVar = MyBoolVar
for c in ['X','Y','W','E','N','S','NW','NE','SW','SE','CENTER','LEFT','RIGHT','TOP',
          'BOTTOM','BOTH','NONE','END','ACTIVE','DISABLED','NORMAL','HORIZONTAL',
          'VERTICAL','SINGLE','EXTENDED','RAISED','SUNKEN','GROOVE','RIDGE','FLAT']:
    setattr(tk_mod, c, c)
tk_mod.font = make_mod(); tk_mod.font.Font = Fake
# ttk
ttk_mod = make_mod()
for cls in ['Treeview','Style','Notebook','Combobox','Frame','Label','Button',
            'Scrollbar','Progressbar','Separator','PanedWindow']:
    setattr(ttk_mod, cls, Fake)
tk_mod.ttk = ttk_mod
# messagebox / filedialog
mb = make_mod()
for m in ['showerror','showwarning','showinfo','showquestion']:
    setattr(mb, m, lambda *a, **k: None)
for m in ['askyesno','askokcancel','askyesnocancel','askretrycancel']:
    setattr(mb, m, lambda *a, **k: True)
fd = make_mod()
fd.asksaveasfilename = lambda *a, **k: os.path.join(tempfile.gettempdir(), 'selftest_out.xlsx')
fd.askopenfilename = lambda *a, **k: ''
tk_mod.messagebox = mb
tk_mod.filedialog = fd
# openpyxl
ox = make_mod(); ox.Workbook = Fake
oxs = make_mod()
for n in ['Font','Alignment','Border','Side','PatternFill','fills','numbers']:
    setattr(oxs, n, Fake)
ox.styles = oxs
# openpyxl.utils 子模块（源码运行时调用 get_column_letter，需返回字符串作 dict 键）
oxu = make_mod()
oxu.get_column_letter = lambda n: 'A'
ox.utils = oxu
sys.modules['openpyxl.utils'] = oxu

sys.modules['tkinter'] = tk_mod
sys.modules['tkinter.ttk'] = ttk_mod
sys.modules['tkinter.font'] = tk_mod.font
sys.modules['tkinter.messagebox'] = mb
sys.modules['tkinter.filedialog'] = fd
sys.modules['openpyxl'] = ox
sys.modules['openpyxl.styles'] = oxs

results = []
def check(name, fn):
    try:
        fn(); results.append(('PASS', name))
    except Exception as e:
        results.append(('FAIL', name, traceback.format_exc()))

import todo_list_v5 as appmod

def t_import():
    assert appmod.THEME['colors']['title_bar'] == '#2C3E50', '默认 title_bar 应为稳重风 #2C3E50'
    assert appmod.LOCK_PORT == 54321, '默认端口应为 54321'
    assert appmod.THEME['icon'] == 'TodoList_icon.ico'

def t_instantiate():
    root = Fake()
    app = appmod.TodoApp(root)
    assert app is not None

def t_save_load():
    root = Fake()
    app = appmod.TodoApp(root)
    app.data_file = os.path.join(tempfile.gettempdir(), 'selftest_todo.json')
    app.save_data()
    app.load_data()

def t_export():
    root = Fake()
    app = appmod.TodoApp(root)
    app.data_file = os.path.join(tempfile.gettempdir(), 'selftest_todo2.json')
    app.save_data()
    app.export_completed()

def t_theme_override():
    # 验证 load_theme 能合并覆盖（不破坏默认）
    ov = {'colors': {'title_bar': '#FF0000'}, 'port': 55555, 'icon': 'X.ico'}
    import json
    p = os.path.join(tempfile.gettempdir(), 'theme_test.json')
    with open(p, 'w', encoding='utf-8') as f: json.dump(ov, f)
    t = appmod._deep_merge(appmod.copy.deepcopy(appmod.DEFAULT_THEME), ov)
    assert t['colors']['title_bar'] == '#FF0000'
    assert t['colors']['primary'] == '#4F81BD'  # 未覆盖保留默认
    assert t['port'] == 55555

check('引擎导入+THEME默认加载', t_import)
check('实例化 TodoApp(无异常)', t_instantiate)
check('save_data/load_data', t_save_load)
check('export_completed', t_export)
check('主题深合并覆盖逻辑', t_theme_override)

print('==== 自测汇总 ====')
ok = sum(1 for r in results if r[0]=='PASS')
for r in results:
    if r[0]=='PASS': print('[PASS]', r[1])
    else: print('[FAIL]', r[1]); print(r[2])
print(f'通过 {ok}/{len(results)}')
