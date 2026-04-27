#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import ctypes
import os
import re
import sys
import time
import datetime as dt
import threading
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import cv2
import numpy as np
from PIL import ImageGrab
from pynput import keyboard
import win32api
import win32con
import win32gui
import win32process
import psutil


WINDOW_KEYWORD = "异环"
TARGET_PROCESS_NAMES = {"HTGame.exe"}
TARGET_WIDTH = 1600
TARGET_HEIGHT = 900
POLL_INTERVAL_SEC = 0.1
CLICK_INTERVAL_SEC = 0.1
PRELOAD_DELAY_SEC = 0.4
PRELOAD_BLOCK_RADIUS = 120
ENABLE_PRELOAD = False
GLOBAL_START_SILENT_SEC = 10.0
GLOBAL_START_TRIGGER_DELAY_SEC = 4.0
MATCH_THRESHOLD = 0.82
CAKE_THRESHOLD_BONUS = 0.05
RED_COLLAR_THRESHOLD = 0.92
# 速度优先：减少缩放档位，降低单轮匹配次数
COARSE_TEMPLATE_SCALES = (
    0.65, 1.00, 1.05, 1.15
)
TEMPLATE_SCALES = (
    0.65, 0.70, 0.75, 0.80, 0.85, 0.90, 0.95, 1.00, 1.05, 1.10, 1.15
)
# 识别ROI（基于客户端相对坐标）
ROI_LEFT = 290
ROI_TOP = 100
ROI_RIGHT = 1260
ROI_BOTTOM = 400

COORDS_MD_NAME = "坐标记录_手动补充_2026-04-25.md"
ASSETS_DIR_NAME = "素材"


def resolve_resource_root() -> Path:
    """
    资源根目录：
    - PyInstaller 单文件：资源打包在 datas 中，运行期解压到 sys._MEIPASS。
    - 源码运行：脚本所在目录。
    """
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = resolve_resource_root()
COORDS_MD_PATH = BASE_DIR / "数据记录" / "配置" / COORDS_MD_NAME
LEGACY_COORDS_MD_PATH = BASE_DIR / COORDS_MD_NAME
ASSETS_DIR = BASE_DIR / ASSETS_DIR_NAME
PLASTIC_TEMPLATE_PATHS = [
    ASSETS_DIR / "塑料杯咖啡.png",
    ASSETS_DIR / "塑料杯咖啡2.png",
]
CERAMIC_TEMPLATE_PATHS = [
    ASSETS_DIR / "瓷杯咖啡.png",
    ASSETS_DIR / "瓷杯咖啡2.png",
]
CROISSANT_TEMPLATE_PATHS = [
    ASSETS_DIR / "牛角包三明治.png",
]
BREAD_TEMPLATE_PATHS = [
    ASSETS_DIR / "面包三明治.png",
]
CAKE_TEMPLATE_PATHS = [
    ASSETS_DIR / "小蛋糕.png",
]
RED_COLLAR_TEMPLATE_PATHS = [
    ASSETS_DIR / "红领子.png",
]
GAME_START_TEMPLATE_PATHS = [
    ASSETS_DIR / "游戏开始.png",
]


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


def parse_coords_from_md(md_path: Path) -> Dict[str, Tuple[int, int]]:
    text = md_path.read_text(encoding="utf-8")
    # 兼容中英文冒号、可选反引号，避免格式轻微变化导致字段漏读
    pattern = re.compile(r"-\s*([^:：\n]+)\s*[:：]\s*`?\((\d+),\s*(\d+)\)`?")
    coords: Dict[str, Tuple[int, int]] = {}
    for name, x, y in pattern.findall(text):
        coords[name.strip()] = (int(x), int(y))
    return coords


def resolve_coords_md_path() -> Path:
    """
    优先使用新配置目录；兼容旧版根目录坐标文件。
    """
    if COORDS_MD_PATH.exists():
        return COORDS_MD_PATH
    return LEGACY_COORDS_MD_PATH


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


def screenshot_client_bgr(hwnd: int) -> Optional[np.ndarray]:
    try:
        (ox, oy), (w, h) = get_client_origin_and_size(hwnd)
        if w <= 0 or h <= 0:
            return None
        img = ImageGrab.grab(bbox=(ox, oy, ox + w, oy + h))
        arr = np.array(img)
        return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    except Exception:
        return None


def crop_detection_roi(frame_bgr: np.ndarray) -> Optional[np.ndarray]:
    h, w = frame_bgr.shape[:2]
    left = max(0, min(ROI_LEFT, w))
    top = max(0, min(ROI_TOP, h))
    right = max(0, min(ROI_RIGHT, w))
    bottom = max(0, min(ROI_BOTTOM, h))
    if right <= left or bottom <= top:
        return None
    return frame_bgr[top:bottom, left:right]


def match_template_score(screen_gray: np.ndarray, template_gray: np.ndarray) -> float:
    if template_gray is None:
        return 0.0
    sh, sw = screen_gray.shape[:2]
    th, tw = template_gray.shape[:2]
    if sh < th or sw < tw:
        return 0.0
    res = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
    _min_val, max_val, _min_loc, _max_loc = cv2.minMaxLoc(res)
    return float(max_val)


def max_match_score(screen_gray: np.ndarray, templates: List[np.ndarray]) -> float:
    if not templates:
        return 0.0
    return max(match_template_score(screen_gray, tpl) for tpl in templates)


def max_match_score_with_center(screen_gray: np.ndarray, templates: List[np.ndarray]) -> Tuple[float, Optional[Tuple[int, int]]]:
    if not templates:
        return 0.0, None

    best_score = 0.0
    best_center: Optional[Tuple[int, int]] = None
    sh, sw = screen_gray.shape[:2]

    for tpl in templates:
        th, tw = tpl.shape[:2]
        if sh < th or sw < tw:
            continue
        res = cv2.matchTemplate(screen_gray, tpl, cv2.TM_CCOEFF_NORMED)
        _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(res)
        score = float(max_val)
        if score > best_score:
            cx = max_loc[0] + tw // 2
            cy = max_loc[1] + th // 2
            best_score = score
            best_center = (cx, cy)

    return best_score, best_center


def build_scaled_templates(base_templates: List[np.ndarray], scales: Tuple[float, ...]) -> List[np.ndarray]:
    scaled_templates: List[np.ndarray] = []
    for tpl in base_templates:
        th, tw = tpl.shape[:2]
        for scale in scales:
            nw = max(8, int(round(tw * scale)))
            nh = max(8, int(round(th * scale)))
            # 1.0 尺寸保持原图，避免重复插值损失
            if nw == tw and nh == th:
                gray = cv2.cvtColor(tpl, cv2.COLOR_BGR2GRAY)
                scaled_templates.append(gray)
                continue
            resized = cv2.resize(tpl, (nw, nh), interpolation=cv2.INTER_AREA if scale < 1.0 else cv2.INTER_LINEAR)
            gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
            scaled_templates.append(gray)
    return scaled_templates


def imread_unicode(path: Path) -> Optional[np.ndarray]:
    """
    Windows 下 cv2.imread 对中文路径可能失败，这里改用 fromfile + imdecode。
    """
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except Exception:
        return None


def click_abs(x: int, y: int) -> None:
    # 避免偶发越界坐标导致 SetCursorPos 失败
    vs_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    vs_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    vs_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    vs_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    vs_right = vs_left + max(1, vs_width) - 1
    vs_bottom = vs_top + max(1, vs_height) - 1

    clamped_x = min(max(int(x), vs_left), vs_right)
    clamped_y = min(max(int(y), vs_top), vs_bottom)
    if clamped_x != x or clamped_y != y:
        now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(
            f"[{now}] 警告: 点击坐标越界，已钳制 "
            f"({x}, {y}) -> ({clamped_x}, {clamped_y})，"
            f"虚拟屏幕=({vs_left},{vs_top})-({vs_right},{vs_bottom})"
        )

    win32api.SetCursorPos((clamped_x, clamped_y))
    time.sleep(0.03)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def click_rel_postmessage(hwnd: int, rel_x: int, rel_y: int) -> None:
    # 降级方案：直接向窗口客户区投递鼠标消息，规避 SetCursorPos 在部分系统环境报错
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
        raise RuntimeError(
            f"相对坐标超出客户区: ({rel_x}, {rel_y}), client=({cw}, {ch})"
        )
    (ox, oy), _size = get_client_origin_and_size(hwnd)
    abs_x, abs_y = ox + rel_x, oy + rel_y
    now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    print(f"[{now}] 点击 -> 相对({rel_x}, {rel_y}) 绝对({abs_x}, {abs_y})")
    try:
        click_abs(abs_x, abs_y)
    except Exception as exc:
        now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[{now}] SetCursorPos 点击失败，改用 PostMessage 降级点击: {exc}")
        click_rel_postmessage(hwnd, rel_x, rel_y)


def run_plastic_workflow(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    for key in ("塑料杯", "咖啡", "刮花"):
        x, y = coords[key]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)


def run_ceramic_workflow(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    for key in ("瓷杯", "咖啡", "牛奶"):
        x, y = coords[key]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)


def run_refill_coffee_click(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    x, y = coords["补充咖啡"]
    click_rel(hwnd, x, y)
    time.sleep(CLICK_INTERVAL_SEC)


def run_croissant_workflow(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    for key in ("切牛角包", "拿牛角包", "鸡蛋配料"):
        x, y = coords[key]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)
    # 按需求：函数末尾再执行一次“切牛角包”
    x, y = coords["切牛角包"]
    click_rel(hwnd, x, y)
    time.sleep(CLICK_INTERVAL_SEC)


def run_bread_workflow(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    for key in ("切面包", "拿面包", "培根配料"):
        x, y = coords[key]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)
    # 按需求：函数末尾再执行一次“切面包”
    x, y = coords["切面包"]
    click_rel(hwnd, x, y)
    time.sleep(CLICK_INTERVAL_SEC)


def run_cake_workflow(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    for key in ("烘焙蛋糕", "拿蛋糕", "果酱"):
        x, y = coords[key]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)
    # 按需求：函数末尾再执行一次“烘焙蛋糕”
    x, y = coords["烘焙蛋糕"]
    click_rel(hwnd, x, y)
    time.sleep(CLICK_INTERVAL_SEC)


def run_game_start_global_sequence(hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    """
    全局监测触发流程：
    补充咖啡 -> 切牛角包 -> 切面包 -> 烘焙蛋糕，然后静默等待。
    """
    print("触发全局流程：补充咖啡 -> 切牛角包 -> 切面包 -> 烘焙蛋糕")
    for key in ("补充咖啡", "切牛角包", "切面包", "烘焙蛋糕"):
        x, y = coords[key]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)


def print_similarity_snapshot(
    tpl_plastic_list: List[np.ndarray],
    tpl_ceramic_list: List[np.ndarray],
    tpl_croissant_list: List[np.ndarray],
    tpl_bread_list: List[np.ndarray],
    tpl_cake_list: List[np.ndarray],
    tpl_red_collar_list: List[np.ndarray],
    tpl_game_start_list: List[np.ndarray],
) -> None:
    hwnd = pick_game_window()
    if hwnd is None:
        print("[F12] 未找到异环窗口，无法计算相似度。")
        return

    frame = screenshot_client_bgr(hwnd)
    if frame is None:
        print("[F12] 截图失败，无法计算相似度。")
        return
    roi_bgr = crop_detection_roi(frame)
    if roi_bgr is None:
        print("[F12] ROI 无效，无法计算相似度。")
        return
    frame_gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)

    pid = get_window_pid(hwnd)
    exe = get_process_name(pid)
    title = win32gui.GetWindowText(hwnd)
    _origin, size = get_client_origin_and_size(hwnd)
    plastic_score = max_match_score(frame_gray, tpl_plastic_list)
    ceramic_score = max_match_score(frame_gray, tpl_ceramic_list)
    croissant_score = max_match_score(frame_gray, tpl_croissant_list)
    bread_score = max_match_score(frame_gray, tpl_bread_list)
    cake_score = max_match_score(frame_gray, tpl_cake_list)
    red_collar_score = max_match_score(frame_gray, tpl_red_collar_list)
    game_start_score = max_match_score(frame_gray, tpl_game_start_list)
    now = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]

    print(
        f"[{now}] [F12] 相似度快照 -> plastic={plastic_score:.4f}, ceramic={ceramic_score:.4f}, "
        f"croissant={croissant_score:.4f}, bread={bread_score:.4f}, cake={cake_score:.4f}, "
        f"red_collar={red_collar_score:.4f}, game_start={game_start_score:.4f}, "
        f"threshold={MATCH_THRESHOLD:.4f}"
    )
    print(f"[F12] 命中窗口 -> exe={exe} pid={pid} size={size[0]}x{size[1]} title={title}")


def choose_best_action(
    frame_gray: np.ndarray,
    tpl_plastic_list: List[np.ndarray],
    tpl_ceramic_list: List[np.ndarray],
    tpl_croissant_list: List[np.ndarray],
    tpl_bread_list: List[np.ndarray],
    tpl_cake_list: List[np.ndarray],
    tpl_red_collar_list: List[np.ndarray],
    excluded_types: Optional[set[str]] = None,
    blocked_center: Optional[Tuple[int, int]] = None,
    blocked_radius: int = 0,
) -> Tuple[Optional[str], Dict[str, float], Dict[str, Optional[Tuple[int, int]]]]:
    scored = {
        "plastic": max_match_score_with_center(frame_gray, tpl_plastic_list),
        "ceramic": max_match_score_with_center(frame_gray, tpl_ceramic_list),
        "croissant": max_match_score_with_center(frame_gray, tpl_croissant_list),
        "bread": max_match_score_with_center(frame_gray, tpl_bread_list),
        "cake": max_match_score_with_center(frame_gray, tpl_cake_list),
        "hammer": max_match_score_with_center(frame_gray, tpl_red_collar_list),
    }
    scores = {k: v[0] for k, v in scored.items()}
    centers = {k: v[1] for k, v in scored.items()}

    def is_blocked(center: Optional[Tuple[int, int]]) -> bool:
        if blocked_center is None or center is None or blocked_radius <= 0:
            return False
        dx = center[0] - blocked_center[0]
        dy = center[1] - blocked_center[1]
        return dx * dx + dy * dy <= blocked_radius * blocked_radius

    def threshold_for(action_type: str) -> float:
        if action_type == "cake":
            return MATCH_THRESHOLD + CAKE_THRESHOLD_BONUS
        return MATCH_THRESHOLD

    candidates = [
        (k, v) for k, v in scores.items()
        if (not excluded_types or k not in excluded_types)
        and not is_blocked(centers.get(k))
        and v >= threshold_for(k)
    ]
    if not candidates:
        return None, scores, centers
    best_type, best_score = max(candidates, key=lambda item: item[1])
    return best_type, scores, centers


def run_action_once(best_type: str, best_score: float, hwnd: int, coords: Dict[str, Tuple[int, int]]) -> None:
    if best_type == "plastic":
        print(f"识别到塑料杯咖啡模板，score={best_score:.3f}，执行：塑料杯 -> 咖啡 -> 刮花")
        print("执行前置步骤：补充咖啡")
        run_refill_coffee_click(hwnd, coords)
        run_plastic_workflow(hwnd, coords)
        print("执行后置步骤：补充咖啡")
        run_refill_coffee_click(hwnd, coords)
    elif best_type == "ceramic":
        print(f"识别到瓷杯咖啡模板，score={best_score:.3f}，执行：瓷杯 -> 咖啡 -> 牛奶")
        print("执行前置步骤：补充咖啡")
        run_refill_coffee_click(hwnd, coords)
        run_ceramic_workflow(hwnd, coords)
        print("执行后置步骤：补充咖啡")
        run_refill_coffee_click(hwnd, coords)
    elif best_type == "croissant":
        print(f"识别到牛角包三明治模板，score={best_score:.3f}，执行第一组流程。")
        run_croissant_workflow(hwnd, coords)
    elif best_type == "bread":
        print(f"识别到面包三明治模板，score={best_score:.3f}，执行第二组流程。")
        run_bread_workflow(hwnd, coords)
    elif best_type == "hammer":
        print(f"识别到红领子模板，score={best_score:.3f}，执行：点击大锤。")
        x, y = coords["大锤"]
        click_rel(hwnd, x, y)
        time.sleep(CLICK_INTERVAL_SEC)
    else:
        print(f"识别到小蛋糕模板，score={best_score:.3f}，执行第三组流程。")
        run_cake_workflow(hwnd, coords)


def main() -> None:
    dpi_mode = enable_dpi_awareness()
    print(f"DPI 感知模式: {dpi_mode}")

    coords_md_path = resolve_coords_md_path()
    if not coords_md_path.exists():
        raise FileNotFoundError(f"坐标文件不存在: {COORDS_MD_PATH}")
    for p in (
        PLASTIC_TEMPLATE_PATHS
        + CERAMIC_TEMPLATE_PATHS
        + CROISSANT_TEMPLATE_PATHS
        + BREAD_TEMPLATE_PATHS
        + CAKE_TEMPLATE_PATHS
        + RED_COLLAR_TEMPLATE_PATHS
        + GAME_START_TEMPLATE_PATHS
    ):
        if not p.exists():
            raise FileNotFoundError(f"模板不存在: {p}")

    coords = parse_coords_from_md(coords_md_path)
    required_keys = {
        "塑料杯", "瓷杯", "咖啡", "刮花", "牛奶", "补充咖啡",
        "切牛角包", "拿牛角包", "鸡蛋配料",
        "切面包", "拿面包", "培根配料",
        "烘焙蛋糕", "拿蛋糕", "果酱",
        "大锤",
    }
    missing = required_keys - set(coords.keys())
    if missing:
        raise ValueError(f"坐标文件缺少字段: {sorted(missing)}")

    base_tpl_plastic = [imread_unicode(p) for p in PLASTIC_TEMPLATE_PATHS]
    base_tpl_ceramic = [imread_unicode(p) for p in CERAMIC_TEMPLATE_PATHS]
    base_tpl_croissant = [imread_unicode(p) for p in CROISSANT_TEMPLATE_PATHS]
    base_tpl_bread = [imread_unicode(p) for p in BREAD_TEMPLATE_PATHS]
    base_tpl_cake = [imread_unicode(p) for p in CAKE_TEMPLATE_PATHS]
    base_tpl_red_collar = [imread_unicode(p) for p in RED_COLLAR_TEMPLATE_PATHS]
    base_tpl_game_start = [imread_unicode(p) for p in GAME_START_TEMPLATE_PATHS]
    if any(
        tpl is None
        for tpl in (
            base_tpl_plastic
            + base_tpl_ceramic
            + base_tpl_croissant
            + base_tpl_bread
            + base_tpl_cake
            + base_tpl_red_collar
            + base_tpl_game_start
        )
    ):
        raise ValueError("模板图片读取失败，请确认图片未损坏。")
    tpl_plastic_coarse = build_scaled_templates(base_tpl_plastic, COARSE_TEMPLATE_SCALES)
    tpl_ceramic_coarse = build_scaled_templates(base_tpl_ceramic, COARSE_TEMPLATE_SCALES)
    tpl_croissant_coarse = build_scaled_templates(base_tpl_croissant, COARSE_TEMPLATE_SCALES)
    tpl_bread_coarse = build_scaled_templates(base_tpl_bread, COARSE_TEMPLATE_SCALES)
    tpl_cake_coarse = build_scaled_templates(base_tpl_cake, COARSE_TEMPLATE_SCALES)
    tpl_red_collar_coarse = build_scaled_templates(base_tpl_red_collar, COARSE_TEMPLATE_SCALES)

    tpl_plastic_list = build_scaled_templates(base_tpl_plastic, TEMPLATE_SCALES)
    tpl_ceramic_list = build_scaled_templates(base_tpl_ceramic, TEMPLATE_SCALES)
    tpl_croissant_list = build_scaled_templates(base_tpl_croissant, TEMPLATE_SCALES)
    tpl_bread_list = build_scaled_templates(base_tpl_bread, TEMPLATE_SCALES)
    tpl_cake_list = build_scaled_templates(base_tpl_cake, TEMPLATE_SCALES)
    tpl_red_collar_list = build_scaled_templates(base_tpl_red_collar, TEMPLATE_SCALES)
    tpl_game_start_list = build_scaled_templates(base_tpl_game_start, TEMPLATE_SCALES)

    print("已启动图片捕捉制作脚本。按 Ctrl + C 退出。")
    print("提醒：当前坐标按 1600x900 基准运行。")
    print("如需适配全部16:9分辨率，请先确认是否启用归一化。")
    print("按 F12 可输出当前模板组相似度与阈值。")
    print(f"粗匹配比例: {COARSE_TEMPLATE_SCALES}")
    print(f"回退匹配比例: {TEMPLATE_SCALES}")

    def on_press(key: keyboard.KeyCode) -> None:
        if key == keyboard.Key.f12:
            try:
                print_similarity_snapshot(
                    tpl_plastic_list,
                    tpl_ceramic_list,
                    tpl_croissant_list,
                    tpl_bread_list,
                    tpl_cake_list,
                    tpl_red_collar_list,
                    tpl_game_start_list,
                )
            except Exception as exc:
                print(f"[F12] 相似度快照失败: {exc}")
        elif key == keyboard.Key.f10:
            print("[F10] 收到结束指令，正在退出进程。")
            os._exit(0)

    hotkey_listener = keyboard.Listener(on_press=on_press)
    hotkey_listener.daemon = True
    hotkey_listener.start()

    preloaded_action: Optional[str] = None
    preloaded_scores: Optional[Dict[str, float]] = None
    game_start_cooldown_until = 0.0
    preload_thread: Optional[threading.Thread] = None
    preload_holder: Dict[str, object] = {}

    while True:
        # 回收后台预载结果（跨轮保留），避免算出来却丢失
        if ENABLE_PRELOAD and preload_thread is not None and not preload_thread.is_alive():
            next_action = preload_holder.get("action")
            next_scores = preload_holder.get("scores")
            if isinstance(next_action, str) and isinstance(next_scores, dict):
                preloaded_action = next_action
                preloaded_scores = next_scores
                print(f"预载完成：next={next_action}, score={next_scores.get(next_action, 0.0):.3f}")
            preload_thread = None
            preload_holder = {}

        hwnd = pick_game_window()
        if hwnd is None:
            print("未找到异环窗口，等待 0.5s ...")
            time.sleep(POLL_INTERVAL_SEC)
            continue

        frame = screenshot_client_bgr(hwnd)
        if frame is None:
            time.sleep(POLL_INTERVAL_SEC)
            continue
        roi_bgr = crop_detection_roi(frame)
        if roi_bgr is None:
            time.sleep(POLL_INTERVAL_SEC)
            continue
        frame_gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)

        # 全局监测仅对“游戏开始”做冷却，不影响其他模板捕捉
        now_ts = time.time()
        if now_ts >= game_start_cooldown_until:
            game_start_score = max_match_score(frame_gray, tpl_game_start_list)
            if game_start_score >= MATCH_THRESHOLD:
                print(f"识别到游戏开始模板，score={game_start_score:.3f}")
                print(f"等待 {GLOBAL_START_TRIGGER_DELAY_SEC:.1f}s 后执行全局点击流程")
                time.sleep(GLOBAL_START_TRIGGER_DELAY_SEC)
                run_game_start_global_sequence(hwnd, coords)
                game_start_cooldown_until = time.time() + GLOBAL_START_SILENT_SEC
                print(f"游戏开始模板进入静默 {GLOBAL_START_SILENT_SEC:.1f}s（仅暂停该模板检测）")
                preloaded_action = None
                preloaded_scores = None
                preload_thread = None
                preload_holder = {}

        # 红领子最高优先级：每轮先独立匹配，命中则不再走其他模板
        red_collar_score = max_match_score(frame_gray, tpl_red_collar_list)
        if red_collar_score >= RED_COLLAR_THRESHOLD:
            print(f"识别到红领子模板，score={red_collar_score:.3f}，优先执行大锤。")
            run_action_once("hammer", red_collar_score, hwnd, coords)
            preloaded_action = None
            preloaded_scores = None
            preload_thread = None
            preload_holder = {}
            time.sleep(POLL_INTERVAL_SEC)
            continue

        current_hit_center: Optional[Tuple[int, int]] = None
        if ENABLE_PRELOAD and preloaded_action is not None and preloaded_scores is not None:
            best_type = preloaded_action
            best_score = preloaded_scores[best_type]
            print(f"命中预载动作：{best_type}，score={best_score:.3f}")
            preloaded_action = None
            preloaded_scores = None
        else:
            best_type, scores, centers = choose_best_action(
                frame_gray,
                tpl_plastic_coarse,
                tpl_ceramic_coarse,
                tpl_croissant_coarse,
                tpl_bread_coarse,
                tpl_cake_coarse,
                tpl_red_collar_coarse,
            )
            if best_type is None:
                best_type, scores, centers = choose_best_action(
                    frame_gray,
                tpl_plastic_list,
                tpl_ceramic_list,
                tpl_croissant_list,
                tpl_bread_list,
                tpl_cake_list,
                    tpl_red_collar_list,
                )
            if best_type is None:
                time.sleep(POLL_INTERVAL_SEC)
                continue
            best_score = scores[best_type]
            current_hit_center = centers.get(best_type)

        def preload_next_action_worker() -> None:
            time.sleep(PRELOAD_DELAY_SEC)
            next_frame = screenshot_client_bgr(hwnd)
            if next_frame is None:
                return
            next_roi_bgr = crop_detection_roi(next_frame)
            if next_roi_bgr is None:
                return
            next_gray = cv2.cvtColor(next_roi_bgr, cv2.COLOR_BGR2GRAY)
            next_action, next_scores, _next_centers = choose_best_action(
                next_gray,
                tpl_plastic_coarse,
                tpl_ceramic_coarse,
                tpl_croissant_coarse,
                tpl_bread_coarse,
                tpl_cake_coarse,
                tpl_red_collar_coarse,
                blocked_center=current_hit_center,
                blocked_radius=PRELOAD_BLOCK_RADIUS,
            )
            if next_action is None:
                next_action, next_scores, _next_centers = choose_best_action(
                    next_gray,
                tpl_plastic_list,
                tpl_ceramic_list,
                tpl_croissant_list,
                tpl_bread_list,
                tpl_cake_list,
                    tpl_red_collar_list,
                blocked_center=current_hit_center,
                blocked_radius=PRELOAD_BLOCK_RADIUS,
                )
            if next_action is not None:
                preload_holder["action"] = next_action
                preload_holder["scores"] = next_scores

        # 新预载时机：函数开始后0.5s抓帧，预测下一轮动作
        # 新预载时机：函数开始后0.5s抓帧，预测下一轮动作（单轮预载）
        if ENABLE_PRELOAD and preload_thread is None:
            preload_thread = threading.Thread(target=preload_next_action_worker, daemon=True)
            preload_thread.start()

        run_action_once(best_type, best_score, hwnd, coords)

        time.sleep(POLL_INTERVAL_SEC)


if __name__ == "__main__":
    main()
