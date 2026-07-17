# -*- coding: utf-8 -*-
"""
TodoList 多套主题 · 一键批量打包器

设计（单引擎 + 主题配置）：
- engine/todo_list_v5.py   : 唯一源码真相源（已主题化，运行时读 theme.json）
- themes/<套名>/theme.json : 该套主题配置（name/title/icon/port + 可选 colors 覆盖）
- themes/<套名>/*.ico      : 该套应用图标
- 打包后 exe 落在 themes/<套名>/TodoList_<套名>_<版本号>.exe（文件名带版本号，便于区分新旧；版本号取自 theme.json 的 version 字段，缺省回退引擎 VERSION）
- 各套数据文件 todo_data.json 落在各自 exe 同目录，天然独立

用法（在 venv 中）：
    python build_all.py            # 打包 themes/ 下所有套
    python build_all.py 冒险双剑   # 只打包指定套
"""
import sys
import os
import re
import json
import shutil
import subprocess

BUILD_ROOT = os.path.dirname(os.path.abspath(__file__))
ENGINE = os.path.join(BUILD_ROOT, "engine", "todo_list_v5.py")
THEMES_DIR = os.path.join(BUILD_ROOT, "themes")
WORK_DIR = os.path.join(BUILD_ROOT, ".build")


def find_pyinstaller():
    # 优先用当前解释器（应在含 PyInstaller 的 venv 中）
    return [sys.executable, "-m", "PyInstaller"]


def engine_version():
    """从 engine 源码解析 VERSION 常量，作为 theme 未声明 version 时的回退。"""
    try:
        with open(ENGINE, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip().startswith("VERSION"):
                    m = re.search(r'VERSION\s*=\s*["\']([^"\']+)["\']', line)
                    if m:
                        return m.group(1)
    except Exception:
        pass
    return "V0.0.0"


def build_one(theme_dir):
    tj = os.path.join(theme_dir, "theme.json")
    if not os.path.exists(tj):
        print(f"[跳过] {theme_dir} 缺少 theme.json")
        return False
    with open(tj, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    name = cfg.get("name") or os.path.basename(theme_dir)
    icon = cfg.get("icon", "TodoList_icon.ico")
    icon_path = os.path.join(theme_dir, icon)
    if not os.path.exists(icon_path):
        print(f"[错误] {theme_dir} 缺少图标文件 {icon}")
        return False

    version = cfg.get("version") or engine_version()
    exe_name = f"TodoList_{name}_{version}"
    print(f"\n===== 打包 [{name}] =====")
    print(f"  icon : {icon_path}")
    print(f"  port : {cfg.get('port', '(默认)')}")
    print(f"  title: {cfg.get('title', '(默认)')}")

    os.makedirs(WORK_DIR, exist_ok=True)
    cmd = [
        *find_pyinstaller(),
        "--onefile",
        "--windowed",
        "--name", exe_name,
        "--icon", icon_path,
        "--add-data", f"{tj};.",
        "--add-data", f"{icon_path};.",
    ]
    # 标题栏图标（如配置 title_icon）—— 一并内嵌
    title_icon = cfg.get("title_icon")
    if title_icon:
        title_icon_path = os.path.join(theme_dir, title_icon)
        if os.path.exists(title_icon_path):
            cmd.append("--add-data")
            cmd.append(f"{title_icon_path};.")
            print(f"  title_icon: {title_icon_path}")
        else:
            print(f"[警告] {theme_dir} 缺少标题栏图标 {title_icon}（将回退纯文字）")
    cmd += [
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.font",
        "--hidden-import", "tkinter.messagebox",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "openpyxl",
        "--distpath", theme_dir,
        "--workpath", WORK_DIR,
        "--specpath", WORK_DIR,
        ENGINE,
    ]
    # Windows 上 subprocess 列表参数无需手动加引号（含中文/空格安全）
    rc = subprocess.run(cmd).returncode
    if rc != 0:
        print(f"[失败] [{name}] PyInstaller 返回 {rc}")
        return False
    out = os.path.join(theme_dir, exe_name + ".exe")
    if os.path.exists(out):
        print(f"[成功] {out}")
        return True
    print(f"[失败] 未找到产物 {out}")
    return False


def main():
    only = sys.argv[1:] if len(sys.argv) > 1 else None
    targets = []
    for d in sorted(os.listdir(THEMES_DIR)):
        full = os.path.join(THEMES_DIR, d)
        if not os.path.isdir(full):
            continue
        if only and d not in only:
            continue
        targets.append(full)
    if not targets:
        print("没有可打包的主题（检查 themes/ 目录）。")
        sys.exit(1)
    ok = 0
    for t in targets:
        if build_one(t):
            ok += 1
    print(f"\n===== 完成：{ok}/{len(targets)} 套成功 =====")
    # 清理中间产物
    if os.path.isdir(WORK_DIR):
        shutil.rmtree(WORK_DIR, ignore_errors=True)
    if ok < len(targets):
        sys.exit(1)


if __name__ == "__main__":
    main()
