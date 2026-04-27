#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Noki 风格后台鼠标点击最小测试（纯 PostMessage，单句柄）

用途：
- 对照 test_background_click_minimal.py
- 排除 SendMessage/子窗口坐标换算等干扰因素

说明：
- 坐标为目标窗口客户端坐标（client x/y）
- 仅发送到一个 hwnd，不向 child hwnd 转发
"""

from __future__ import annotations

import argparse
import ctypes
import time
from pathlib import Path
from typing import Optional

import cv2
import mss
import numpy as np
import win32con
import win32gui

WM_ACTIVATE = 0x0006
WA_ACTIVE = 1
TARGET_HWND = 198764
FIXED_CLICK_X = 1100
FIXED_CLICK_Y = 790
TEMPLATE_PATH = Path("素材/钓鱼/开始钓鱼.png")
MATCH_THRESHOLD = 0.80


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def fake_activate_window(hwnd: int) -> None:
    # 保持与 Noki 示例一致：只发送 WM_ACTIVATE
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
) -> tuple[Optional[int], Optional[int], Optional[float], bool]:
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
    client_x = int(match_x) + int(match_w) // 2
    client_y = int(match_y) + int(match_h) // 2
    return client_x, client_y, float(max_val), matched


def click_once_noki_style(hwnd: int, client_x: int, client_y: int, hold_sec: float) -> None:
    lparam = make_lparam_from_client_xy(client_x, client_y)
    fake_activate_window(hwnd)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
    time.sleep(hold_sec)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="异环后台点击最小测试（Noki 风格）")
    parser.add_argument("--hwnd", type=int, default=TARGET_HWND, help="目标窗口句柄（默认写死为 198764）。")
    parser.add_argument("--title-keyword", type=str, default="异环", help="自动查找窗口时使用的标题关键字。")
    parser.add_argument("--x", type=int, default=FIXED_CLICK_X, help="客户端坐标 x（默认写死为 1340）。")
    parser.add_argument("--y", type=int, default=FIXED_CLICK_Y, help="客户端坐标 y（默认写死为 780）。")
    parser.add_argument("--count", type=int, default=1, help="点击次数。")
    parser.add_argument("--interval", type=float, default=0.2, help="两次点击间隔秒数。")
    parser.add_argument("--hold", type=float, default=0.08, help="按下到抬起的间隔秒数。")
    parser.add_argument("--template", type=str, default=str(TEMPLATE_PATH), help="模板图片路径。")
    parser.add_argument("--threshold", type=float, default=MATCH_THRESHOLD, help="模板匹配阈值。")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    hwnd = args.hwnd if args.hwnd is not None else TARGET_HWND
    if not win32gui.IsWindow(hwnd):
        # 写死句柄失效时，兜底按标题关键字找窗口
        hwnd = find_window_by_title_keyword(args.title_keyword)

    if hwnd is None:
        print(f"[ERROR] 未找到目标窗口，title_keyword={args.title_keyword!r}")
        return 1
    if not win32gui.IsWindow(hwnd):
        print(f"[ERROR] hwnd={hwnd} 无效（IsWindow=False）")
        return 1

    title = win32gui.GetWindowText(hwnd) or ""
    template_path = Path(args.template)
    threshold = float(args.threshold)
    match_x, match_y, max_score, matched = match_template_center_client_xy(hwnd, template_path, threshold)
    if matched and match_x is not None and match_y is not None:
        click_x, click_y = match_x, match_y
        print(
            f"[INFO] 模板匹配成功：template={template_path}, max_score={max_score:.4f}, "
            f"client_x={click_x}, client_y={click_y}"
        )
    else:
        click_x, click_y = args.x, args.y
        print(
            f"[WARN] 模板匹配失败：template={template_path}, exists={template_path.exists()}, "
            f"threshold={threshold}, max_score={max_score}"
        )
        print(f"[INFO] 使用兜底坐标：client_x={click_x}, client_y={click_y}")

    print(f"[INFO] target hwnd={hwnd}, title={title!r}")
    print(f"[INFO] click at client=({click_x}, {click_y}), count={args.count}, interval={args.interval}, hold={args.hold}")
    print(f"[INFO] script admin={is_admin()}")

    for i in range(1, max(1, args.count) + 1):
        if not win32gui.IsWindow(hwnd):
            print("[ERROR] 目标窗口已失效，停止发送")
            return 1
        click_once_noki_style(hwnd, click_x, click_y, args.hold)
        print(f"[INFO] sent click #{i}")
        if i < args.count:
            time.sleep(max(0.0, args.interval))

    print("[INFO] done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
