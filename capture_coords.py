#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
基础坐标捕捉脚本（Windows）

功能：
1) 监听全局鼠标右键点击
2) 获取标题包含“异环”的窗口客户端区域（不含标题栏）
3) 以客户端左上角为 (0, 0) 计算相对坐标
4) 按时间写入本地 Markdown 日志
"""

from __future__ import annotations

import ctypes
import datetime as dt
from pathlib import Path
from typing import List, Optional, Tuple

from pynput import mouse
import win32gui
import win32con
import win32process
import psutil


WINDOW_KEYWORD = "异环"
# 通过任务管理器实测可见游戏本体进程名为 HTGame.exe
TARGET_PROCESS_NAMES = {"HTGame.exe"}
TARGET_WIDTH = 1600
TARGET_HEIGHT = 900
LOG_DIR = Path("数据记录/配置/capture_coords")
LOG_FILE = LOG_DIR / f"coords_{dt.date.today().isoformat()}.md"


def enable_dpi_awareness() -> str:
    """
    开启 DPI 感知，避免 Windows 缩放(如150%)导致坐标/尺寸被虚拟化。
    """
    try:
        user32 = ctypes.windll.user32
        shcore = ctypes.windll.shcore

        if hasattr(user32, "SetProcessDpiAwarenessContext"):
            per_monitor_v2 = ctypes.c_void_p(-4)  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            if user32.SetProcessDpiAwarenessContext(per_monitor_v2):
                return "Per-Monitor DPI Aware V2"

        if hasattr(shcore, "SetProcessDpiAwareness"):
            if shcore.SetProcessDpiAwareness(2) == 0:  # PROCESS_PER_MONITOR_DPI_AWARE
                return "Per-Monitor DPI Aware"

        if hasattr(user32, "SetProcessDPIAware") and user32.SetProcessDPIAware():
            return "System DPI Aware"
    except Exception:
        pass

    return "Unknown / Not enabled"


def find_target_window(keyword: str) -> Optional[int]:
    """查找标题包含 keyword 的可见窗口句柄。"""
    handles = []

    def callback(hwnd: int, _extra) -> None:
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if keyword in title:
            handles.append(hwnd)

    win32gui.EnumWindows(callback, None)
    return handles[0] if handles else None


def find_all_target_windows(keyword: str) -> List[int]:
    """查找所有标题包含 keyword 的可见顶层窗口句柄。"""
    handles: List[int] = []

    def callback(hwnd: int, _extra) -> None:
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if keyword in title:
            handles.append(hwnd)

    win32gui.EnumWindows(callback, None)
    return handles


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


def get_top_level_window_from_point(x: int, y: int) -> Optional[int]:
    """从屏幕坐标找到顶层窗口句柄。"""
    hwnd = win32gui.WindowFromPoint((x, y))
    if not hwnd:
        return None
    return win32gui.GetAncestor(hwnd, win32con.GA_ROOT)


def score_size_distance(size: Tuple[int, int]) -> int:
    """计算窗口尺寸与基准分辨率的距离（越小越接近）。"""
    w, h = size
    return abs(w - TARGET_WIDTH) + abs(h - TARGET_HEIGHT)


def pick_best_target_window(click_x: int, click_y: int) -> Optional[int]:
    """
    选择最可能的目标窗口：
    1) 优先点击点所在顶层窗口（标题包含关键字）
    2) 否则在所有候选里按尺寸最接近 1600x900 选择
    """
    clicked_hwnd = get_top_level_window_from_point(click_x, click_y)
    if clicked_hwnd:
        clicked_title = win32gui.GetWindowText(clicked_hwnd)
        clicked_pid = get_window_pid(clicked_hwnd)
        clicked_exe = get_process_name(clicked_pid)
        if clicked_exe in TARGET_PROCESS_NAMES:
            return clicked_hwnd
        if WINDOW_KEYWORD in clicked_title:
            return clicked_hwnd

    candidates = find_all_target_windows(WINDOW_KEYWORD)
    if not candidates:
        return None

    def key_func(hwnd: int) -> Tuple[int, int, int]:
        _origin, size = get_client_origin_and_size(hwnd)
        pid = get_window_pid(hwnd)
        exe = get_process_name(pid)
        # 优先级：游戏本体进程 > 尺寸接近 > hwnd稳定排序
        process_penalty = 0 if exe in TARGET_PROCESS_NAMES else 1
        return process_penalty, score_size_distance(size), hwnd

    candidates.sort(key=key_func)
    return candidates[0]


def get_client_origin_and_size(hwnd: int) -> Tuple[Tuple[int, int], Tuple[int, int]]:
    """
    返回客户端区域：
    - 屏幕坐标下的客户端左上角 (origin_x, origin_y)
    - 客户端尺寸 (width, height)
    """
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width = right - left
    height = bottom - top
    origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
    return (origin_x, origin_y), (width, height)


def init_log_file(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return
    header = (
        f"# 异环坐标记录\n\n"
        f"- 创建时间: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"- 目标分辨率: {TARGET_WIDTH} x {TARGET_HEIGHT}\n"
        f"- 坐标原点: 客户端左上角（不含标题栏）\n\n"
        f"## 记录\n\n"
    )
    path.write_text(header, encoding="utf-8")


def append_log(path: Path, abs_xy: Tuple[int, int], rel_xy: Tuple[int, int], window_size: Tuple[int, int]) -> None:
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = (
        f"- {now} | 绝对坐标: ({abs_xy[0]}, {abs_xy[1]})"
        f" | 相对坐标: ({rel_xy[0]}, {rel_xy[1]})"
        f" | 窗口尺寸: {window_size[0]}x{window_size[1]}\n"
    )
    with path.open("a", encoding="utf-8") as f:
        f.write(line)


def main() -> None:
    dpi_mode = enable_dpi_awareness()
    init_log_file(LOG_FILE)
    print("已启动右键坐标捕捉。按 Ctrl + C 退出。")
    print(f"DPI 感知模式: {dpi_mode}")
    print(f"目标窗口关键字: {WINDOW_KEYWORD}")
    print(f"日志文件: {LOG_FILE.resolve()}")

    def on_click(x: int, y: int, button: mouse.Button, pressed: bool) -> None:
        if not pressed or button != mouse.Button.right:
            return

        try:
            hwnd = pick_best_target_window(x, y)
            if hwnd is None:
                print("未找到标题包含“异环”的窗口，已忽略本次点击。")
                return

            pid = get_window_pid(hwnd)
            exe = get_process_name(pid)
            title = win32gui.GetWindowText(hwnd)
            (origin_x, origin_y), (w, h) = get_client_origin_and_size(hwnd)
            rel_x, rel_y = x - origin_x, y - origin_y

            if not (0 <= rel_x < w and 0 <= rel_y < h):
                print(f"右键不在客户端内：屏幕({x}, {y})，客户端原点({origin_x}, {origin_y})。")
                return

            if (w, h) != (TARGET_WIDTH, TARGET_HEIGHT):
                print(
                    f"警告：当前客户端尺寸为 {w}x{h}，与目标 {TARGET_WIDTH}x{TARGET_HEIGHT} 不一致，仍已记录。"
                )

            append_log(LOG_FILE, (x, y), (rel_x, rel_y), (w, h))
            print(f"命中窗口 -> exe={exe} pid={pid} title={title}")
            print(f"已记录 -> 相对坐标: ({rel_x}, {rel_y})")
        except Exception as exc:
            # 保底保护：单次点击异常不应导致全局监听退出
            print(f"捕捉失败（已跳过本次点击）: {exc}")

    with mouse.Listener(on_click=on_click) as listener:
        listener.join()


if __name__ == "__main__":
    main()
