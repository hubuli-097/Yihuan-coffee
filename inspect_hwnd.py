#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
inspect_hwnd.py

用于排查《异环》相关窗口句柄（hwnd）信息。

依赖：
    pip install pywin32 psutil

功能：
1) 打印当前前台窗口信息
2) 枚举所有可见顶层窗口，筛选：
   - 标题包含“异环”
   - 或进程名为 HTGame.exe
3) 打印每个匹配顶层窗口的详细信息
4) 枚举并打印该顶层窗口的所有子窗口详细信息

注意：
- 本脚本只读取和打印信息，不会发送按键，不会修改窗口状态。
"""

from __future__ import annotations

import argparse
import sys
from typing import Dict, List, Optional, Tuple

import psutil
import win32gui
import win32process


def safe_get_window_text(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def safe_get_class_name(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def safe_get_window_rect(hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    try:
        return win32gui.GetWindowRect(hwnd)
    except Exception:
        return None


def safe_get_client_rect(hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    try:
        return win32gui.GetClientRect(hwnd)
    except Exception:
        return None


def safe_client_to_screen(hwnd: int, point: Tuple[int, int]) -> Optional[Tuple[int, int]]:
    try:
        return win32gui.ClientToScreen(hwnd, point)
    except Exception:
        return None


def safe_get_pid(hwnd: int) -> Optional[int]:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid if pid > 0 else None
    except Exception:
        return None


def safe_get_exe_name(pid: Optional[int], cache: Dict[int, str]) -> str:
    if not pid:
        return "<unknown>"
    if pid in cache:
        return cache[pid]
    try:
        name = psutil.Process(pid).name()
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        name = "<unknown>"
    except Exception:
        name = "<unknown>"
    cache[pid] = name
    return name


def is_target_window(title: str, exe_name: str) -> bool:
    return ("异环" in title) or (exe_name.lower() == "htgame.exe")


def print_kv(key: str, value: object, indent: int = 0) -> None:
    prefix = " " * indent
    print(f"{prefix}{key}: {value}")


def format_rect(rect: Optional[Tuple[int, int, int, int]]) -> str:
    if rect is None:
        return "<unavailable>"
    l, t, r, b = rect
    return f"({l}, {t}, {r}, {b})  w={r - l}, h={b - t}"


def format_point(pt: Optional[Tuple[int, int]]) -> str:
    if pt is None:
        return "<unavailable>"
    return f"({pt[0]}, {pt[1]})"


def print_window_info(
    hwnd: int,
    exe_cache: Dict[int, str],
    fg_hwnd: Optional[int],
    indent: int = 0,
    prefix: str = "",
) -> None:
    title = safe_get_window_text(hwnd)
    cls = safe_get_class_name(hwnd)
    pid = safe_get_pid(hwnd)
    exe_name = safe_get_exe_name(pid, exe_cache)
    wrect = safe_get_window_rect(hwnd)
    crect = safe_get_client_rect(hwnd)
    client_origin = safe_client_to_screen(hwnd, (0, 0))

    if prefix:
        print(" " * indent + prefix)
    print_kv("hwnd", hwnd, indent)
    print_kv("title", repr(title), indent)
    print_kv("class name", cls, indent)
    print_kv("pid", pid, indent)
    print_kv("exe name", exe_name, indent)
    print_kv("window rect", format_rect(wrect), indent)
    print_kv("client rect", format_rect(crect), indent)
    print_kv("client origin screen", format_point(client_origin), indent)
    if fg_hwnd is not None:
        print_kv("is foreground", hwnd == fg_hwnd, indent)


def enum_visible_top_windows() -> List[int]:
    windows: List[int] = []

    def _cb(hwnd: int, _: object) -> bool:
        if win32gui.IsWindowVisible(hwnd):
            windows.append(hwnd)
        return True

    win32gui.EnumWindows(_cb, None)
    return windows


def enum_child_windows(parent_hwnd: int) -> List[int]:
    children: List[int] = []

    def _cb(hwnd: int, _: object) -> bool:
        children.append(hwnd)
        return True

    win32gui.EnumChildWindows(parent_hwnd, _cb, None)
    return children


def print_foreground_window(exe_cache: Dict[int, str]) -> Optional[int]:
    print("=" * 88)
    print("当前前台窗口")
    print("=" * 88)
    try:
        fg_hwnd = win32gui.GetForegroundWindow()
    except Exception:
        fg_hwnd = 0
    if not fg_hwnd:
        print("foreground hwnd: <none>")
        print()
        return None

    title = safe_get_window_text(fg_hwnd)
    cls = safe_get_class_name(fg_hwnd)
    pid = safe_get_pid(fg_hwnd)
    exe_name = safe_get_exe_name(pid, exe_cache)

    print_kv("foreground hwnd", fg_hwnd)
    print_kv("title", repr(title))
    print_kv("class name", cls)
    print_kv("pid", pid)
    print_kv("exe name", exe_name)
    print()
    return fg_hwnd


def inspect_target_windows(show_all_children: bool = True) -> int:
    exe_cache: Dict[int, str] = {}
    fg_hwnd = print_foreground_window(exe_cache)
    top_windows = enum_visible_top_windows()

    matched_top: List[int] = []
    for hwnd in top_windows:
        title = safe_get_window_text(hwnd)
        pid = safe_get_pid(hwnd)
        exe_name = safe_get_exe_name(pid, exe_cache)
        if is_target_window(title, exe_name):
            matched_top.append(hwnd)

    print("=" * 88)
    print("匹配到的顶层窗口")
    print("=" * 88)
    print(f"visible top windows total: {len(top_windows)}")
    print(f"matched top windows: {len(matched_top)}")
    print()

    if not matched_top:
        print("未匹配到标题包含“异环”或进程名为 HTGame.exe 的可见顶层窗口。")
        return 0

    for idx, hwnd in enumerate(matched_top, start=1):
        print("-" * 88)
        print(f"[Top {idx}/{len(matched_top)}]")
        print_window_info(hwnd, exe_cache, fg_hwnd, indent=2)
        print()

        children = enum_child_windows(hwnd)
        print("  " + "." * 80)
        print(f"  子窗口数量: {len(children)}")
        print("  " + "." * 80)

        if not children:
            print("  <无子窗口>")
            print()
            continue

        for c_idx, ch in enumerate(children, start=1):
            title = safe_get_window_text(ch)
            cls = safe_get_class_name(ch)
            wrect = safe_get_window_rect(ch)
            crect = safe_get_client_rect(ch)
            corigin = safe_client_to_screen(ch, (0, 0))
            visible = False
            enabled = False
            try:
                visible = bool(win32gui.IsWindowVisible(ch))
            except Exception:
                pass
            try:
                enabled = bool(win32gui.IsWindowEnabled(ch))
            except Exception:
                pass

            print(f"    [Child {c_idx}/{len(children)}]")
            print_kv("child hwnd", ch, 6)
            print_kv("parent hwnd", hwnd, 6)
            print_kv("title", repr(title), 6)
            print_kv("class name", cls, 6)
            print_kv("window rect", format_rect(wrect), 6)
            print_kv("client rect", format_rect(crect), 6)
            print_kv("client origin screen", format_point(corigin), 6)
            print_kv("is visible", visible, 6)
            print_kv("is enabled", enabled, 6)
            print()

            if not show_all_children:
                break

    return len(matched_top)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="排查《异环》窗口句柄（hwnd）：枚举匹配顶层窗口及其子窗口。"
    )
    parser.add_argument(
        "--first-child-only",
        action="store_true",
        help="仅打印每个匹配顶层窗口的第一个子窗口（默认打印全部子窗口）",
    )
    args = parser.parse_args()

    print("inspect_hwnd.py - 只读检查模式（不会发送按键/不会修改窗口状态）")
    print()

    try:
        matched_count = inspect_target_windows(show_all_children=not args.first_child_only)
    except KeyboardInterrupt:
        print("\n用户中断。")
        return 130
    except Exception as exc:
        print(f"\n执行失败: {exc.__class__.__name__}: {exc}")
        return 1

    print("=" * 88)
    print(f"完成。匹配到顶层窗口数量: {matched_count}")
    print("=" * 88)
    return 0


if __name__ == "__main__":
    sys.exit(main())
