#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
最小 F 发送测试（窗口消息版）
目标：验证后台状态下，是否能把 F 发到异环窗口（仅 PostMessage）。

运行：
    ./.venv/Scripts/python.exe test_send_f_minimal.py

注意：
    本脚本固定使用 TARGET_HWND，不自动找窗口，不切前台，不激活窗口。
"""

from __future__ import annotations

import time
from typing import Optional

import psutil
import win32con
import win32gui
import win32process


TARGET_HWND = 198764
PRESS_COUNT = 5
PRESS_INTERVAL_SEC = 0.2
KEYDOWN_HOLD_SEC = 0.1
WM_ACTIVATE = 0x0006
WA_ACTIVE = 1


def get_window_pid(hwnd: int) -> Optional[int]:
    try:
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return None


def get_process_name(pid: Optional[int]) -> str:
    if not pid:
        return "unknown"
    try:
        return psutil.Process(pid).name()
    except Exception:
        return "unknown"


def get_exe_name_by_hwnd(hwnd: int) -> str:
    pid = get_window_pid(hwnd)
    return get_process_name(pid)


def get_class_name(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def get_window_rect(hwnd: int) -> tuple[int, int, int, int] | str:
    try:
        return win32gui.GetWindowRect(hwnd)
    except Exception:
        return "unavailable"


def get_client_rect(hwnd: int) -> tuple[int, int, int, int] | str:
    try:
        return win32gui.GetClientRect(hwnd)
    except Exception:
        return "unavailable"


def get_client_origin_screen(hwnd: int) -> tuple[int, int] | str:
    try:
        return win32gui.ClientToScreen(hwnd, (0, 0))
    except Exception:
        return "unavailable"


def send_key_postmessage_simple(hwnd: int, vk: int, hold_sec: float) -> None:
    """
    Noki 风格最简 PostMessage 发送：
    PostMessage(hwnd, WM_KEYDOWN, vk, 0)
    sleep(hold_sec)
    PostMessage(hwnd, WM_KEYUP, vk, 0)
    """
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
    time.sleep(hold_sec)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)


def fake_activate_window(hwnd: int) -> None:
    win32gui.SendMessage(hwnd, WM_ACTIVATE, WA_ACTIVE, 0)


def print_target_window_info(hwnd: int) -> None:
    title = win32gui.GetWindowText(hwnd) or ""
    class_name = get_class_name(hwnd)
    pid = get_window_pid(hwnd)
    exe = get_process_name(pid)
    rect = get_window_rect(hwnd)
    client_rect = get_client_rect(hwnd)
    client_origin_screen = get_client_origin_screen(hwnd)
    visible = win32gui.IsWindowVisible(hwnd)
    enabled = win32gui.IsWindowEnabled(hwnd)

    print("[INFO] TARGET_HWND 信息：")
    print(f"  - hwnd: {hwnd}")
    print(f"  - title: {title}")
    print(f"  - class name: {class_name}")
    print(f"  - pid: {pid}")
    print(f"  - exe name: {exe}")
    print(f"  - window rect: {rect}")
    print(f"  - client rect: {client_rect}")
    print(f"  - client origin screen: {client_origin_screen}")
    print(f"  - visible: {visible}")
    print(f"  - enabled: {enabled}")


def print_foreground_info() -> int:
    fg_hwnd = win32gui.GetForegroundWindow()
    fg_title = win32gui.GetWindowText(fg_hwnd) if fg_hwnd else ""
    fg_class = get_class_name(fg_hwnd) if fg_hwnd else ""
    fg_pid = get_window_pid(fg_hwnd) if fg_hwnd else None
    fg_exe = get_process_name(fg_pid)
    print("[INFO] 当前前台窗口信息：")
    print(f"  - foreground hwnd: {fg_hwnd}")
    print(f"  - foreground title: {fg_title}")
    print(f"  - foreground class: {fg_class}")
    print(f"  - foreground exe: {fg_exe}")
    return fg_hwnd


def main() -> None:
    hwnd = TARGET_HWND
    if not win32gui.IsWindow(hwnd):
        print(f"[ERROR] TARGET_HWND={hwnd} 无效（IsWindow=False）。")
        print("[ERROR] 请重新运行 inspect_hwnd.py 获取新的 hwnd 后再测。")
        return

    print_target_window_info(hwnd)
    fg_hwnd = print_foreground_info()
    is_fg = hwnd == fg_hwnd
    print(f"[INFO] TARGET_HWND 是否是 foreground: {is_fg}")
    if not is_fg:
        print("[INFO] TARGET_HWND 不是 foreground，仍继续发送 F（后台测试模式）。")

    print(f"[INFO] 开始发送 F：count={PRESS_COUNT}, interval={PRESS_INTERVAL_SEC}s")

    try:
        for i in range(1, PRESS_COUNT + 1):
            if not win32gui.IsWindow(hwnd):
                print("[ERROR] 目标窗口失效，停止发送。")
                break
            fake_activate_window(hwnd)
            print("fake activate sent")
            print(f"send F #{i}")
            send_key_postmessage_simple(hwnd, 0x46, KEYDOWN_HOLD_SEC)
            time.sleep(PRESS_INTERVAL_SEC)
    except KeyboardInterrupt:
        print("\n[INFO] 收到 Ctrl+C，已停止。")

    print("[INFO] 测试结束。")


if __name__ == "__main__":
    main()
