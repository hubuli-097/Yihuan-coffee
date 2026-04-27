#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
最小后台点击测试（窗口消息版）

目标：
- 独立验证异环后台点击链路是否能推动游戏状态
- 固定点击客户端坐标 (1340, 780)

运行：
    ./.venv/Scripts/python.exe test_background_click_minimal.py --hwnd <目标窗口句柄>

说明：
- 坐标必须是客户端坐标（非屏幕坐标、非 ROI 局部坐标）
- 消息顺序严格为：
  WM_MOUSEMOVE -> sleep(0.05) -> WM_LBUTTONDOWN -> sleep(0.08) -> WM_LBUTTONUP -> sleep(0.05) -> WM_MOUSEMOVE
"""

from __future__ import annotations

import argparse
import ctypes
import sys
import time
from pathlib import Path
from typing import Optional, Tuple

import cv2
import mss
import numpy as np
import win32con
import win32gui


WINDOW_TITLE_KEYWORD = "异环"
TEMPLATE_PATH = Path("素材/钓鱼/开始钓鱼.png")
MATCH_THRESHOLD = 0.80
FALLBACK_CLICK_X = 800
FALLBACK_CLICK_Y = 400
WM_ACTIVATE = 0x0006
WA_ACTIVE = 1
WM_SETFOCUS = 0x0007


def enable_dpi_awareness() -> str:
    user32 = ctypes.windll.user32
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        ok = user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        if ok:
            return "PerMonitorV2"
    except Exception:
        pass
    try:
        shcore = ctypes.windll.shcore
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        shcore.SetProcessDpiAwareness(2)
        return "PerMonitor"
    except Exception:
        pass
    try:
        user32.SetProcessDPIAware()
        return "SystemAware"
    except Exception:
        return "Unaware"


def fake_activate_window(hwnd: int) -> None:
    win32gui.SendMessage(hwnd, WM_SETFOCUS, 0, 0)
    win32gui.SendMessage(hwnd, WM_ACTIVATE, WA_ACTIVE, 0)


def make_lparam_from_client_xy(client_x: int, client_y: int) -> int:
    return (int(client_y) << 16) | (int(client_x) & 0xFFFF)


def find_window_by_title_keyword(keyword: str) -> Optional[int]:
    found: list[int] = []

    def _enum_cb(hwnd: int, _lparam: object) -> bool:
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd) or ""
        if keyword in title:
            found.append(hwnd)
        return True

    win32gui.EnumWindows(_enum_cb, None)
    return found[0] if found else None


def print_click_debug_info(hwnd: int, client_x: int, client_y: int, lparam: int, did_fake_activate: bool) -> None:
    fg_hwnd = win32gui.GetForegroundWindow()
    print("[CLICK_DEBUG]")
    print(f"  hwnd: {hwnd}")
    print(f"  client_x: {client_x}")
    print(f"  client_y: {client_y}")
    print(f"  lParam: {lparam} (0x{lparam:08X})")
    print(f"  fake_activate: {did_fake_activate}")
    print(f"  foreground_hwnd: {fg_hwnd}")


def capture_client_bgr(hwnd: int) -> Optional[np.ndarray]:
    try:
        left_top = win32gui.ClientToScreen(hwnd, (0, 0))
        _l, _t, width, height = win32gui.GetClientRect(hwnd)
    except win32gui.error:
        return None
    if width <= 0 or height <= 0:
        return None
    left, top = left_top
    monitor = {"left": left, "top": top, "width": width, "height": height}
    with mss.MSS() as sct:
        shot = sct.grab(monitor)
    bgra = np.asarray(shot, dtype=np.uint8)
    return cv2.cvtColor(bgra, cv2.COLOR_BGRA2BGR)


def load_template(path: Path) -> Optional[np.ndarray]:
    if not path.exists():
        return None
    data = np.fromfile(str(path), dtype=np.uint8)
    if data.size == 0:
        return None
    return cv2.imdecode(data, cv2.IMREAD_COLOR)


def match_template_center_client_xy(
    hwnd: int, template_path: Path, threshold: float
) -> Tuple[Optional[int], Optional[int], Optional[float], bool]:
    frame = capture_client_bgr(hwnd)
    if frame is None:
        return None, None, None, False
    template = load_template(template_path)
    if template is None:
        return None, None, None, False
    if frame.shape[0] < template.shape[0] or frame.shape[1] < template.shape[1]:
        return None, None, None, False

    result = cv2.matchTemplate(frame, template, cv2.TM_CCOEFF_NORMED)
    _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(result)
    matched = bool(max_val >= threshold)

    match_x, match_y = max_loc
    match_h, match_w = template.shape[:2]
    roi_left = 0
    roi_top = 0
    # 若来自模板匹配，按 ROI 局部坐标换算到客户端坐标并取中心点
    client_x = roi_left + int(match_x) + int(match_w) // 2
    client_y = roi_top + int(match_y) + int(match_h) // 2
    return client_x, client_y, float(max_val), matched


def background_left_click_minimal(hwnd: int, client_x: int, client_y: int) -> None:
    # 1) 点击前 fake_activate_window(hwnd)
    fake_activate_window(hwnd)
    did_fake_activate = True

    # 4) lParam = (y << 16) | (x & 0xFFFF)
    lparam = make_lparam_from_client_xy(client_x, client_y)
    print_click_debug_info(hwnd, client_x, client_y, lparam, did_fake_activate)
    child_hwnd = win32gui.ChildWindowFromPoint(hwnd, (client_x, client_y))
    if child_hwnd and child_hwnd != hwnd:
        print(f"[CLICK_DEBUG] child_hwnd_at_point: {child_hwnd}")
    else:
        child_hwnd = None

    # 2) 完整鼠标事件链路（并按 5/6/7 的 wParam 要求）
    # 先 Post 一条 MOVE，给窗口更新鼠标热点
    win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
    time.sleep(0.05)

    # 使用 SendMessage 同步发送 DOWN/UP，降低只收到按下不收到松开的概率
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    if child_hwnd is not None:
        win32gui.SendMessage(child_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    time.sleep(0.08)

    # WM_LBUTTONUP, wParam = 0
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    if child_hwnd is not None:
        win32gui.SendMessage(child_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    # 部分游戏窗口偶发吞掉首次 UP，补发一次兜底
    time.sleep(0.02)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    if child_hwnd is not None:
        win32gui.PostMessage(child_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(0.02)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    if child_hwnd is not None:
        win32gui.PostMessage(child_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(0.05)

    # 尾部补一条 WM_MOUSEMOVE, wParam = 0
    win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="异环最小后台点击测试脚本")
    parser.add_argument(
        "--hwnd",
        type=int,
        default=None,
        help="目标窗口 hwnd。不传则自动按标题关键字“异环”查找第一个可见顶层窗口。",
    )
    parser.add_argument(
        "--template",
        type=str,
        default=str(TEMPLATE_PATH),
        help="模板图片路径（默认 素材/钓鱼/开始钓鱼.png）",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=MATCH_THRESHOLD,
        help="模板匹配阈值（默认 0.80）",
    )
    return parser.parse_args()


def main() -> int:
    dpi_mode = enable_dpi_awareness()
    print(f"[INFO] DPI awareness: {dpi_mode}")

    args = parse_args()
    hwnd = args.hwnd if args.hwnd is not None else find_window_by_title_keyword(WINDOW_TITLE_KEYWORD)

    if hwnd is None:
        print("[ERROR] 未找到目标窗口。请传 --hwnd，或确保窗口标题包含“异环”。")
        return 1
    if not win32gui.IsWindow(hwnd):
        print(f"[ERROR] hwnd={hwnd} 无效（IsWindow=False）。")
        return 1

    template_path = Path(args.template)
    threshold = float(args.threshold)
    client_x, client_y, max_score, matched = match_template_center_client_xy(hwnd, template_path, threshold)
    if not matched or client_x is None or client_y is None:
        print(
            f"[ERROR] 模板匹配失败：template={template_path}, exists={template_path.exists()}, "
            f"threshold={threshold}, max_score={max_score}"
        )
        print(
            f"[INFO] 使用兜底坐标执行后台点击：hwnd={hwnd}, "
            f"client_x={FALLBACK_CLICK_X}, client_y={FALLBACK_CLICK_Y}"
        )
        background_left_click_minimal(hwnd, FALLBACK_CLICK_X, FALLBACK_CLICK_Y)
        print("[INFO] 兜底点击消息已发送。")
        return 1
    print(
        f"[INFO] 模板匹配成功：template={template_path}, max_score={max_score:.4f}, "
        f"client_x={client_x}, client_y={client_y}"
    )
    print(f"[INFO] 将执行后台点击：hwnd={hwnd}, client_x={client_x}, client_y={client_y}")
    background_left_click_minimal(hwnd, client_x, client_y)
    print("[INFO] 点击消息已发送。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
