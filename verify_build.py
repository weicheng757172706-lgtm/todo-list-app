# -*- coding: utf-8 -*-
"""验证各套打包产物：exe 存在 + 内嵌 theme.json + 图标资源。

原理：PyInstaller --onefile 把 --add-data 注入的文件打进 exe 内的 _MEIPASS 包，
其文件名（theme.json / *.ico）与 ico 签名会作为字节留在 exe 中，可被静态检索。

用法（在 venv 中）：
    python verify_build.py
"""
import os
import re
import sys
import json

ROOT = os.path.dirname(os.path.abspath(__file__))
THEMES = os.path.join(ROOT, "themes")
ENGINE = os.path.join(ROOT, "engine", "todo_list_v5.py")


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


def verify(theme_dir):
    tj = os.path.join(theme_dir, "theme.json")
    with open(tj, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    name = cfg.get("name") or os.path.basename(theme_dir)
    icon = cfg.get("icon", "TodoList_icon.ico")
    version = cfg.get("version") or engine_version()
    exe = os.path.join(theme_dir, f"TodoList_{name}_{version}.exe")
    print(f"\n--- [{name}] ---")
    if not os.path.exists(exe):
        print("[FAIL] exe 不存在:", exe)
        return False
    size = os.path.getsize(exe)
    print(f"[OK] exe 存在 ({size/1024/1024:.1f} MB)")
    with open(exe, "rb") as f:
        blob = f.read()
    checks = {
        "内嵌 theme.json": b"theme.json" in blob,
        f"内嵌图标文件名 {icon}": icon.encode("utf-8") in blob,
        "ico 签名 (00 00 01 00)": b"\x00\x00\x01\x00" in blob,
    }
    title_icon = cfg.get("title_icon")
    if title_icon:
        checks[f"内嵌标题栏图标 {title_icon}"] = title_icon.encode("utf-8") in blob
        checks["PNG 签名 (89 50 4E 47)"] = b"\x89PNG" in blob
    else:
        print("  [SKIP] 该套未配置 title_icon（双剑套保持纯文字）")
    ok = True
    for k, v in checks.items():
        print(f"  [{'OK' if v else 'FAIL'}] {k}")
        ok = ok and v
    return ok


def main():
    results = []
    for d in sorted(os.listdir(THEMES)):
        full = os.path.join(THEMES, d)
        if os.path.isdir(full) and os.path.exists(os.path.join(full, "theme.json")):
            results.append(verify(full))
    passed = sum(results)
    print(f"\n===== 验证：{passed}/{len(results)} 套通过 =====")
    sys.exit(0 if (results and all(results)) else 1)


if __name__ == "__main__":
    main()
