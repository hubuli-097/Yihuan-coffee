"""
最小单文件钓鱼辅助脚本（异环，固定 1600x900 客户端坐标）

依赖安装：
    pip install pywin32 mss opencv-python numpy pydirectinput
"""

from __future__ import annotations

import os
import time
import ctypes
from pathlib import Path
from typing import Optional, Tuple

import cv2
import mss
import numpy as np
import pydirectinput
import win32gui
import win32process
import psutil
import win32con
import win32ui


# =========================
# 可调参数（集中放顶部）
# =========================
WINDOW_TITLE_KEYWORD = "异环"
TARGET_PROCESS_NAMES = {"HTGame.exe"}

ROI_LEFT = 500
ROI_TOP = 50
ROI_RIGHT = 1100
ROI_BOTTOM = 70

# HSV 阈值（OpenCV: H[0,179], S[0,255], V[0,255]）
# 基于 素材/钓鱼/钓鱼测试.png 自动测算值做“稍微放宽”：
#   green base: [76,158,114] ~ [92,233,255]
#   yellow base:[19,45,89]  ~ [34,189,255]
HSV_GREEN_LOWER = np.array([72, 135, 95], dtype=np.uint8)
HSV_GREEN_UPPER = np.array([96, 255, 255], dtype=np.uint8)

YELLOW_HSV_PROFILES = {
    # 严格：按黄色样本 5~95 分位数附近，误检更少
    "strict": (
        np.array([25, 68, 232], dtype=np.uint8),
        np.array([30, 133, 255], dtype=np.uint8),
    ),
    # 宽松：适度放宽，提升波动场景下检出率
    "loose": (
        np.array([24, 60, 225], dtype=np.uint8),
        np.array([32, 150, 255], dtype=np.uint8),
    ),
}
YELLOW_HSV_PROFILE_NAME = os.environ.get("FISHING_YELLOW_HSV_PROFILE", "strict").strip().lower()
if YELLOW_HSV_PROFILE_NAME not in YELLOW_HSV_PROFILES:
    YELLOW_HSV_PROFILE_NAME = "strict"
HSV_YELLOW_LOWER, HSV_YELLOW_UPPER = YELLOW_HSV_PROFILES[YELLOW_HSV_PROFILE_NAME]

DEAD_ZONE = 10
LOOP_INTERVAL = 0.04
DEBUG = True
START_DELAY_ENABLED = False
START_DELAY_SECONDS = 0.0
USE_FOREGROUND_FALLBACK = True
CAPTURE_BACKEND = "printwindow"  # "mss" | "printwindow"
INPUT_BACKEND = "postmessage"  # "pydirectinput" | "postmessage"
ALLOW_MSS_FALLBACK_WHEN_PW_FAIL = False
STARTUP_STABLE_FRAMES = 2
WAIT_FOR_START_SIGNAL = os.environ.get("FISHING_BOT_WAIT_SIGNAL", "0") == "1"
START_SIGNAL_PATH_RAW = os.environ.get("FISHING_BOT_START_SIGNAL_PATH", "").strip()

# 调试参数
DEBUG_SAVE_INTERVAL = 0.8
DEBUG_IMAGE_PATH = Path("数据记录/调试/fishing_bot/debug_roi.png")


# 仅固定 1600x900 逻辑（按用户要求不做归一化）
CLIENT_BASE_WIDTH = 1600
CLIENT_BASE_HEIGHT = 900
PW_CLIENTONLY = 0x00000001
PW_RENDERFULLCONTENT = 0x00000002


def enable_dpi_awareness() -> str:
    """
    开启 DPI 感知，避免坐标被系统缩放虚拟化。
    """
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


def find_window_handle_by_keyword(keyword: str) -> Optional[int]:
    """查找标题中包含 keyword 的可见窗口句柄。"""
    keyword_lower = keyword.lower()
    matched_hwnd: Optional[int] = None

    def _enum_proc(hwnd: int, _lparam: int) -> bool:
        nonlocal matched_hwnd
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd) or ""
        if keyword_lower in title.lower():
            matched_hwnd = hwnd
            return False  # 停止枚举
        return True

    try:
        win32gui.EnumWindows(_enum_proc, 0)
    except win32gui.error as exc:
        print(f"[WARN] EnumWindows 失败：{exc}")
        return None
    return matched_hwnd


def get_window_handle(keyword: str) -> Optional[int]:
    """先枚举窗口；若失败则尝试前台窗口兜底。"""
    hwnd = find_window_handle_by_keyword(keyword)
    if hwnd is not None:
        return hwnd

    if not USE_FOREGROUND_FALLBACK:
        return None

    fg_hwnd = win32gui.GetForegroundWindow()
    if not fg_hwnd or not win32gui.IsWindow(fg_hwnd):
        return None

    title = win32gui.GetWindowText(fg_hwnd) or ""
    pid = get_window_pid(fg_hwnd)
    exe = get_process_name(pid)
    if keyword.lower() in title.lower() or exe in TARGET_PROCESS_NAMES:
        print(f"[INFO] 使用前台窗口兜底成功 hwnd={fg_hwnd} title={title} exe={exe} pid={pid}")
        return fg_hwnd
    return None


def get_client_origin_and_size(hwnd: int) -> Optional[Tuple[int, int, int, int]]:
    """
    返回客户端左上角屏幕坐标与客户端大小：
    (client_origin_x, client_origin_y, client_width, client_height)
    """
    if not win32gui.IsWindow(hwnd):
        return None
    try:
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        client_w = right - left
        client_h = bottom - top
        origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
    except win32gui.error:
        return None
    return origin_x, origin_y, client_w, client_h


def capture_roi_bgr(
    sct: mss.mss, client_origin_x: int, client_origin_y: int
) -> np.ndarray:
    """按固定客户端 ROI 抓取图像并返回 BGR。"""
    screen_left = client_origin_x + ROI_LEFT
    screen_top = client_origin_y + ROI_TOP
    width = ROI_RIGHT - ROI_LEFT
    height = ROI_BOTTOM - ROI_TOP

    monitor = {
        "left": screen_left,
        "top": screen_top,
        "width": width,
        "height": height,
    }
    shot = sct.grab(monitor)
    bgra = np.array(shot, dtype=np.uint8)
    bgr = cv2.cvtColor(bgra, cv2.COLOR_BGRA2BGR)
    return bgr


def capture_roi_bgr_by_printwindow(hwnd: int) -> Optional[np.ndarray]:
    """
    使用 PrintWindow 抓取客户端，再裁剪 ROI。
    适合窗口不在最前台但仍可被 GDI 抓取的场景。
    """
    client_info = get_client_origin_and_size(hwnd)
    if client_info is None:
        return None
    _ox, _oy, client_w, client_h = client_info
    if client_w <= 0 or client_h <= 0:
        return None

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    if hwnd_dc == 0:
        return None

    # 不同游戏/渲染路径对 PrintWindow flag 兼容性不同，逐个尝试
    flags_to_try = (
        PW_CLIENTONLY | PW_RENDERFULLCONTENT,
        PW_RENDERFULLCONTENT,
        PW_CLIENTONLY,
        0,
    )

    try:
        for pw_flag in flags_to_try:
            mfc_dc = None
            save_dc = None
            save_bitmap = None
            try:
                mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
                save_dc = mfc_dc.CreateCompatibleDC()
                save_bitmap = win32ui.CreateBitmap()
                save_bitmap.CreateCompatibleBitmap(mfc_dc, client_w, client_h)
                save_dc.SelectObject(save_bitmap)

                result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), pw_flag)
                if result != 1:
                    continue

                bmp_info = save_bitmap.GetInfo()
                bmp_bytes = save_bitmap.GetBitmapBits(True)
                img = np.frombuffer(bmp_bytes, dtype=np.uint8)
                img = img.reshape((bmp_info["bmHeight"], bmp_info["bmWidth"], 4))
                client_bgr = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

                left = max(0, ROI_LEFT)
                top = max(0, ROI_TOP)
                right = min(client_bgr.shape[1], ROI_RIGHT)
                bottom = min(client_bgr.shape[0], ROI_BOTTOM)
                if right <= left or bottom <= top:
                    continue

                roi = client_bgr[top:bottom, left:right]
                # 纯黑/近纯黑通常表示抓图失败；做轻量健康检查
                if float(np.mean(roi)) < 3.0 and float(np.std(roi)) < 3.0:
                    continue
                return roi
            finally:
                if save_bitmap is not None:
                    win32gui.DeleteObject(save_bitmap.GetHandle())
                if save_dc is not None:
                    save_dc.DeleteDC()
                if mfc_dc is not None:
                    mfc_dc.DeleteDC()
        return None
    finally:
        win32gui.ReleaseDC(hwnd, hwnd_dc)


def detect_green_region_client_x(
    roi_bgr: np.ndarray,
) -> Optional[Tuple[int, int, int]]:
    """识别绿色区域，返回客户端坐标 (green_left, green_right, green_mid)。"""
    hsv = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, HSV_GREEN_LOWER, HSV_GREEN_UPPER)

    # 轻微去噪，避免孤立像素误检
    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None

    # 绿色条通常是细长横向区域，优先取面积最大的连通域
    cnt = max(contours, key=cv2.contourArea)
    area = cv2.contourArea(cnt)
    if area < 80:
        return None

    x, _y, w, _h = cv2.boundingRect(cnt)
    if w < 20:
        return None

    green_left_local = int(x)
    green_right_local = int(x + w - 1)
    green_mid_local = int((green_left_local + green_right_local) / 2)

    green_left = ROI_LEFT + green_left_local
    green_right = ROI_LEFT + green_right_local
    green_mid = ROI_LEFT + green_mid_local
    return green_left, green_right, green_mid


def detect_yellow_line_client_x(roi_bgr: np.ndarray) -> Optional[int]:
    """识别黄色竖线，返回客户端坐标 yellow_x。"""
    hsv = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, HSV_YELLOW_LOWER, HSV_YELLOW_UPPER)

    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None

    # 黄色目标是竖线：偏高、偏窄，先过滤再选面积最大
    candidates = []
    for cnt in contours:
        x, _y, w, h = cv2.boundingRect(cnt)
        area = cv2.contourArea(cnt)
        if area < 10:
            continue
        if h < 8:
            continue
        if w > 18:
            continue
        candidates.append((area, x, w))

    if not candidates:
        return None

    _area, x, w = max(candidates, key=lambda item: item[0])
    yellow_local_x = int(x + w / 2)
    return ROI_LEFT + yellow_local_x


def draw_debug_image(
    roi_bgr: np.ndarray,
    green_region: Optional[Tuple[int, int, int]],
    yellow_x_client: Optional[int],
    action: str,
) -> None:
    """保存调试图：标注绿区边界/中心与黄线。"""
    DEBUG_IMAGE_PATH.parent.mkdir(parents=True, exist_ok=True)
    vis = roi_bgr.copy()
    h, _w = vis.shape[:2]

    if green_region is not None:
        green_left, green_right, green_mid = green_region
        gl = green_left - ROI_LEFT
        gr = green_right - ROI_LEFT
        gm = green_mid - ROI_LEFT

        cv2.line(vis, (gl, 0), (gl, h - 1), (0, 255, 0), 2)
        cv2.line(vis, (gr, 0), (gr, h - 1), (0, 200, 0), 2)
        cv2.line(vis, (gm, 0), (gm, h - 1), (255, 200, 0), 2)

    if yellow_x_client is not None:
        yl = yellow_x_client - ROI_LEFT
        cv2.line(vis, (yl, 0), (yl, h - 1), (0, 255, 255), 2)

    text = f"ROI client=({ROI_LEFT},{ROI_TOP})-({ROI_RIGHT},{ROI_BOTTOM}) action={action}"
    cv2.putText(
        vis,
        text,
        (6, h - 8),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.45,
        (255, 255, 255),
        1,
        cv2.LINE_AA,
    )
    cv2.imwrite(str(DEBUG_IMAGE_PATH), vis)


def build_action(green_mid: int, yellow_x: int) -> Tuple[str, int]:
    """根据误差返回动作与误差值。"""
    error = green_mid - yellow_x
    if error > DEAD_ZONE:
        return "D", error
    if error < -DEAD_ZONE:
        return "A", error
    return "NONE", error


last_action = "NONE"
current_hwnd: Optional[int] = None


def send_key_postmessage(hwnd: int, key: str, is_down: bool) -> None:
    vk_map = {"a": 0x41, "d": 0x44}
    vk = vk_map[key]
    if is_down:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
    else:
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0xC0000001)


def set_action(action: str) -> None:
    """A/D/NONE 互斥动作，避免重复 keyDown。"""
    global last_action
    if action == last_action:
        return

    if INPUT_BACKEND == "postmessage":
        if current_hwnd is None or not win32gui.IsWindow(current_hwnd):
            action = "NONE"
        elif action == "A":
            send_key_postmessage(current_hwnd, "a", True)
            send_key_postmessage(current_hwnd, "d", False)
        elif action == "D":
            send_key_postmessage(current_hwnd, "d", True)
            send_key_postmessage(current_hwnd, "a", False)
        else:
            send_key_postmessage(current_hwnd, "a", False)
            send_key_postmessage(current_hwnd, "d", False)
            action = "NONE"
    else:
        if action == "A":
            pydirectinput.keyDown("a")
            pydirectinput.keyUp("d")
        elif action == "D":
            pydirectinput.keyDown("d")
            pydirectinput.keyUp("a")
        else:
            pydirectinput.keyUp("a")
            pydirectinput.keyUp("d")
            action = "NONE"

    last_action = action


def main() -> None:
    dpi_mode = enable_dpi_awareness()
    print(f"[INFO] DPI 感知模式: {dpi_mode}")
    print(f"[INFO] CAPTURE_BACKEND={CAPTURE_BACKEND}, INPUT_BACKEND={INPUT_BACKEND}")
    print(
        f"[INFO] YELLOW_HSV_PROFILE={YELLOW_HSV_PROFILE_NAME} "
        f"lower={HSV_YELLOW_LOWER.tolist()} upper={HSV_YELLOW_UPPER.tolist()}"
    )

    if INPUT_BACKEND == "pydirectinput":
        pydirectinput.PAUSE = 0
        pydirectinput.FAILSAFE = False

    hwnd = get_window_handle(WINDOW_TITLE_KEYWORD)
    if hwnd is None:
        print(
            f"[ERROR] 未找到标题包含“{WINDOW_TITLE_KEYWORD}”的窗口，程序退出。"
            "如 EnumWindows 被拒绝，请先把游戏窗口切到前台再启动脚本。"
        )
        return

    if START_DELAY_ENABLED and START_DELAY_SECONDS > 0:
        print(f"[INFO] 找到窗口 hwnd={hwnd}，{START_DELAY_SECONDS:.1f}s 后开始追踪。")
        time.sleep(START_DELAY_SECONDS)
    else:
        print(f"[INFO] 找到窗口 hwnd={hwnd}，开始钓鱼追踪循环。")

    if WAIT_FOR_START_SIGNAL:
        if not START_SIGNAL_PATH_RAW:
            print("[WARN] 已启用等待启动信号，但未提供信号文件路径，继续直接运行。")
        else:
            signal_path = Path(START_SIGNAL_PATH_RAW)
            print(f"[INFO] 等待启动信号文件：{signal_path}")
            while True:
                if signal_path.exists():
                    print("[INFO] 已收到启动信号，开始钓鱼追踪。")
                    break
                time.sleep(0.02)

    last_debug_save_ts = 0.0

    pw_fallback_warned = False
    startup_valid_count = 0
    with mss.MSS() as sct:
        try:
            while True:
                # 每轮确认窗口存在
                if not win32gui.IsWindow(hwnd):
                    set_action("NONE")
                    print("[WARN] 目标窗口已失效，停止控制。")
                    break

                # 客户端坐标（非窗口外框）
                client_info = get_client_origin_and_size(hwnd)
                if client_info is None:
                    set_action("NONE")
                    print("[WARN] 无法获取客户端坐标，action=NONE")
                    time.sleep(LOOP_INTERVAL)
                    continue

                client_origin_x, client_origin_y, client_w, client_h = client_info
                global current_hwnd
                current_hwnd = hwnd

                # 固定基准提醒（不阻断运行）
                if client_w != CLIENT_BASE_WIDTH or client_h != CLIENT_BASE_HEIGHT:
                    print(
                        "[WARN] 当前客户端分辨率不是 1600x900："
                        f"{client_w}x{client_h}（脚本仍按固定坐标执行）"
                    )

                if CAPTURE_BACKEND == "printwindow":
                    roi_bgr = capture_roi_bgr_by_printwindow(hwnd)
                    if roi_bgr is None:
                        if ALLOW_MSS_FALLBACK_WHEN_PW_FAIL:
                            # 后台抓图失败时，回退 mss（需要窗口可见）
                            if not pw_fallback_warned:
                                print("[WARN] PrintWindow 抓图失败，已回退 mss。窗口被最小化/受保护时常见。")
                                pw_fallback_warned = True
                            roi_bgr = capture_roi_bgr(sct, client_origin_x, client_origin_y)
                        else:
                            set_action("NONE")
                            if not pw_fallback_warned:
                                print("[WARN] PrintWindow 抓图失败（纯黑/无内容），且已禁用 mss 回退。")
                                pw_fallback_warned = True
                            time.sleep(LOOP_INTERVAL)
                            continue
                else:
                    roi_bgr = capture_roi_bgr(sct, client_origin_x, client_origin_y)

                green_region = detect_green_region_client_x(roi_bgr)
                if green_region is None:
                    set_action("NONE")
                    print("green=NONE yellow=UNKNOWN action=NONE")
                    if DEBUG and (time.time() - last_debug_save_ts) >= DEBUG_SAVE_INTERVAL:
                        draw_debug_image(roi_bgr, None, None, "NONE")
                        last_debug_save_ts = time.time()
                    time.sleep(LOOP_INTERVAL)
                    continue

                yellow_x = detect_yellow_line_client_x(roi_bgr)
                if yellow_x is None:
                    set_action("NONE")
                    gl, gr, gm = green_region
                    print(f"green=[{gl},{gr}] mid={gm} yellow=NONE action=NONE")
                    if DEBUG and (time.time() - last_debug_save_ts) >= DEBUG_SAVE_INTERVAL:
                        draw_debug_image(roi_bgr, green_region, None, "NONE")
                        last_debug_save_ts = time.time()
                    time.sleep(LOOP_INTERVAL)
                    continue

                green_left, green_right, green_mid = green_region
                action, error = build_action(green_mid, yellow_x)
                # 启动初期先累计稳定帧，避免第一批误检导致大幅误操作
                if startup_valid_count < STARTUP_STABLE_FRAMES:
                    startup_valid_count += 1
                    set_action("NONE")
                    print(
                        f"green=[{green_left},{green_right}] mid={green_mid} "
                        f"yellow={yellow_x} error={error} action=NONE warmup={startup_valid_count}/{STARTUP_STABLE_FRAMES}"
                    )
                else:
                    set_action(action)
                    print(
                        f"green=[{green_left},{green_right}] mid={green_mid} "
                        f"yellow={yellow_x} error={error} action={action}"
                    )

                if DEBUG and (time.time() - last_debug_save_ts) >= DEBUG_SAVE_INTERVAL:
                    draw_debug_image(roi_bgr, green_region, yellow_x, action)
                    last_debug_save_ts = time.time()

                time.sleep(LOOP_INTERVAL)

        except KeyboardInterrupt:
            print("\n[INFO] 收到 Ctrl+C，准备退出。")
        finally:
            set_action("NONE")
            print("[INFO] 已释放 A / D，脚本结束。")


if __name__ == "__main__":
    main()
