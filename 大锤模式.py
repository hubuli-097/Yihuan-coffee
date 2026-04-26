#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import ctypes
import datetime as dt
import re
import sys
import threading
import time
from pathlib import Path
from typing import Dict, Optional, Tuple

import psutil
import win32api
import win32con
import win32gui
import win32process

WINDOW_KEYWORD = "异环"
TARGET_PROCESS_NAMES = {"HTGame.exe"}
TARGET_WIDTH = 1600
TARGET_HEIGHT = 900
CLICK_INTERVAL_SEC = 0.5
COORDS_MD_NAME = "坐标记录_手动补充_2026-04-25.md"
VK_XBUTTON1 = 0x05  # 鼠标前侧键
VK_XBUTTON2 = 0x06  # 鼠标后侧键


def resolve_resource_root() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = resolve_resource_root()
COORDS_MD_PATH = BASE_DIR / COORDS_MD_NAME


def enable_dpi_awareness() -> str:
    try:
        user32 = ctypes.windll.user32
        shcore = ctypes.windll.shcore

        if hasattr(user32, "SetProcessDpiAwarenessContext"):
            per_monitor_v2 = ctypes.c_void_p(-4)
            if user32.SetProcessDpiAwarenessContext(per_monitor_v2):
                return "Per-Monitor DPI Aware V2"

        if hasattr(shcore, "SetProcessDpiAwareness"):
            if shcore.SetProcessDpiAwareness(2) == 0:
                return "Per-Monitor DPI Aware"

        if hasattr(user32, "SetProcessDPIAware") and user32.SetProcessDPIAware():
            return "System DPI Aware"
    except Exception:
        pass
    return "Unknown / Not enabled"


def parse_coords_from_md(md_path: Path) -> Dict[str, Tuple[int, int]]:
    text = md_path.read_text(encoding="utf-8")
    pattern = re.compile(r"-\s*([^:：\n]+)\s*[:：]\s*`?\((\d+),\s*(\d+)\)`?")
    coords: Dict[str, Tuple[int, int]] = {}
    for name, x, y in pattern.findall(text):
        coords[name.strip()] = (int(x), int(y))
    return coords
def get_window_pid(hwnd: int) -> Optional[int]:
    try:
        _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
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


def get_client_origin_and_size(hwnd: int) -> Tuple[Tuple[int, int], Tuple[int, int]]:
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width = right - left
    height = bottom - top
    origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
    return (origin_x, origin_y), (width, height)


def score_size_distance(size: Tuple[int, int]) -> int:
    w, h = size
    return abs(w - TARGET_WIDTH) + abs(h - TARGET_HEIGHT)


def pick_game_window() -> Optional[int]:
    candidates = []

    def callback(hwnd: int, _extra) -> None:
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if WINDOW_KEYWORD not in title:
            return
        pid = get_window_pid(hwnd)
        exe = get_process_name(pid)
        try:
            _origin, size = get_client_origin_and_size(hwnd)
        except Exception:
            return
        process_penalty = 0 if exe in TARGET_PROCESS_NAMES else 1
        candidates.append((process_penalty, score_size_distance(size), hwnd))

    win32gui.EnumWindows(callback, None)
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    return candidates[0][2]


def click_abs(x: int, y: int) -> None:
    vs_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    vs_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    vs_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    vs_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    vs_right = vs_left + max(1, vs_width) - 1
    vs_bottom = vs_top + max(1, vs_height) - 1

    clamped_x = min(max(int(x), vs_left), vs_right)
    clamped_y = min(max(int(y), vs_top), vs_bottom)
    win32api.SetCursorPos((clamped_x, clamped_y))
    time.sleep(0.03)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def click_rel_postmessage(hwnd: int, rel_x: int, rel_y: int) -> None:
    lparam = (int(rel_y) << 16) | (int(rel_x) & 0xFFFF)
    win32api.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)


def click_rel(hwnd: int, rel_x: int, rel_y: int) -> None:
    if not win32gui.IsWindow(hwnd):
        raise RuntimeError(f"目标窗口句柄无效: hwnd={hwnd}")
    _origin, (cw, ch) = get_client_origin_and_size(hwnd)
    if not (0 <= int(rel_x) < cw and 0 <= int(rel_y) < ch):
        raise RuntimeError(f"相对坐标超出客户区: ({rel_x}, {rel_y}), client=({cw}, {ch})")
    (ox, oy), _size = get_client_origin_and_size(hwnd)
    abs_x, abs_y = ox + rel_x, oy + rel_y
    now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{now}] 点击大锤 -> 相对({rel_x}, {rel_y}) 绝对({abs_x}, {abs_y})")
    try:
        click_abs(abs_x, abs_y)
    except Exception as exc:
        now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[{now}] SetCursorPos 点击失败，改用 PostMessage: {exc}")
        click_rel_postmessage(hwnd, rel_x, rel_y)


def main(auto_start: bool = False, run_gate: threading.Event | None = None) -> None:
    dpi_mode = enable_dpi_awareness()
    print(f"DPI 模式: {dpi_mode}")

    if not COORDS_MD_PATH.exists():
        raise FileNotFoundError(f"未找到坐标文件: {COORDS_MD_PATH}")

    coords = parse_coords_from_md(COORDS_MD_PATH)
    if "大锤" not in coords:
        raise KeyError("坐标文件缺少 '大锤' 坐标。")

    hwnd = pick_game_window()
    if hwnd is None:
        raise RuntimeError("未找到异环窗口，请先打开游戏并保持窗口可见。")

    _origin, size = get_client_origin_and_size(hwnd)
    print(f"已绑定窗口 hwnd={hwnd}, client={size[0]}x{size[1]}")
    print("热键模式：F9/鼠标前侧键 开始，F10/鼠标后侧键 停止，Ctrl+C 退出。")
    print("开始后每 0.5 秒点击一次大锤。")
    if auto_start:
        print("自动模式：启动后立即开始点击大锤。")
    if run_gate is not None:
        print("阻塞模式：仅在主流程放行时点击大锤。")

    x, y = coords["大锤"]
    running = auto_start
    next_click_at = 0.0
    f9_last_down = False
    f10_last_down = False
    x1_last_down = False
    x2_last_down = False

    while True:
        f9_down = bool(win32api.GetAsyncKeyState(win32con.VK_F9) & 0x8000)
        f10_down = bool(win32api.GetAsyncKeyState(win32con.VK_F10) & 0x8000)
        x1_down = bool(win32api.GetAsyncKeyState(VK_XBUTTON1) & 0x8000)
        x2_down = bool(win32api.GetAsyncKeyState(VK_XBUTTON2) & 0x8000)

        # 按下沿触发，避免长按时重复切换状态
        if (f9_down and not f9_last_down) or (x1_down and not x1_last_down):
            running = True
            next_click_at = 0.0
            now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            if x1_down and not x1_last_down:
                print(f"[{now}] 收到鼠标前侧键，开始点击。")
            else:
                print(f"[{now}] 收到 F9，开始点击。")
        if (f10_down and not f10_last_down) or (x2_down and not x2_last_down):
            running = False
            now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            if x2_down and not x2_last_down:
                print(f"[{now}] 收到鼠标后侧键，停止点击。")
            else:
                print(f"[{now}] 收到 F10，停止点击。")

        f9_last_down = f9_down
        f10_last_down = f10_down
        x1_last_down = x1_down
        x2_last_down = x2_down

        if running:
            if run_gate is not None and not run_gate.is_set():
                time.sleep(0.01)
                continue
            now_ts = time.time()
            if now_ts >= next_click_at:
                click_rel(hwnd, x, y)
                next_click_at = now_ts + CLICK_INTERVAL_SEC

        time.sleep(0.01)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已停止大锤模式。")
    except Exception as exc:
        print(f"运行失败: {exc}")
