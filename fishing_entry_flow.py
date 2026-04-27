#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import subprocess
import time
import ctypes
from pathlib import Path
from typing import Optional, Tuple

import cv2
import mss
import numpy as np
import psutil
import win32gui
import win32process
import win32ui
import win32con

from fishing_recovery import BOT_TIMEOUT_SEC, RecoveryState, attempt_fishing_recovery


# ===== 基础配置（固定 1600x900 客户端坐标，不做归一化） =====
WINDOW_TITLE_KEYWORD = "异环"
TARGET_PROCESS_NAMES = {"HTGame.exe"}

CLICK_ROI = (1430, 770, 1530, 870)
CLOSE_ROI = (520, 60, 700, 160)

TEMPLATE_CLICK = Path("素材/钓鱼/点击.png")
TEMPLATE_CLOSE = Path("素材/钓鱼/关闭.png")

POLL_INTERVAL = 0.08
F_PRESS_INTERVAL = 0.18
F_KEYDOWN_HOLD_SEC = 0.1
WM_ACTIVATE = 0x0006
WA_ACTIVE = 1

INPUT_BACKEND = "postmessage"  # "pydirectinput" | "postmessage"

FISHING_BOT_PATH = Path("fishing_bot.py")
PYTHON_EXE = Path(".venv/Scripts/python.exe")
PRELAUNCH_BOT = True

# 截图后端：默认 printwindow，可切换 mss
CAPTURE_BACKEND = "printwindow"  # "printwindow" | "mss"
ALLOW_MSS_FALLBACK_WHEN_PW_FAIL = False
PW_CLIENTONLY = 0x00000001
PW_RENDERFULLCONTENT = 0x00000002
TEMPLATE_SCALES = (1.0,)
# 经验值：alpha+CCORR_NORMED 在这组图标上命中约 0.91
MASK_MATCH_THRESHOLD = 0.88
CLICK_MATCH_THRESHOLD = 0.80
CLOSE_MATCH_THRESHOLD = 0.84
DISAPPEAR_ABSENT_CONSEC_FRAMES = 1
MAX_F_PRESS_COUNT_IN_STAGE2 = 45
CLOSE_HIT_CONSEC_FRAMES = 3


def terminate_bot_process(bot_proc: Optional[subprocess.Popen]) -> None:
    if bot_proc is None:
        return
    if bot_proc.poll() is not None:
        return
    bot_proc.terminate()
    try:
        bot_proc.wait(timeout=3)
    except Exception:
        bot_proc.kill()


def imread_unicode(path: Path) -> np.ndarray | None:
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except Exception:
        return None


def imread_unicode_unchanged(path: Path) -> np.ndarray | None:
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_UNCHANGED)
    except Exception:
        return None


def load_template_with_mask(path: Path) -> tuple[np.ndarray, np.ndarray] | None:
    """
    读取模板与掩码：
    - 优先使用 PNG alpha 作为模板掩码（更稳定）
    - 无 alpha 时回退为白色掩码
    """
    raw = imread_unicode_unchanged(path)
    if raw is None:
        return None

    if len(raw.shape) == 3 and raw.shape[2] == 4:
        bgr = raw[:, :, :3]
        alpha = raw[:, :, 3]
        # 过滤透明边缘噪声
        mask = cv2.inRange(alpha, 20, 255)
    else:
        bgr = raw if len(raw.shape) == 3 else cv2.cvtColor(raw, cv2.COLOR_GRAY2BGR)
        mask = build_white_mask(bgr)

    mask_pixels = int(np.count_nonzero(mask))
    if mask_pixels < 20:
        return None
    return bgr, mask


def load_template_plain(path: Path) -> np.ndarray | None:
    return imread_unicode(path)


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
    matched_hwnd: Optional[int] = None
    keyword_lower = keyword.lower()

    def _enum_proc(hwnd: int, _lparam: int) -> bool:
        nonlocal matched_hwnd
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd) or ""
        if keyword_lower in title.lower():
            matched_hwnd = hwnd
            return False
        return True

    try:
        win32gui.EnumWindows(_enum_proc, 0)
    except win32gui.error as exc:
        print(f"[WARN] EnumWindows 失败：{exc}")
        return None
    return matched_hwnd


def get_window_handle() -> Optional[int]:
    hwnd = find_window_handle_by_keyword(WINDOW_TITLE_KEYWORD)
    if hwnd is not None:
        return hwnd

    fg_hwnd = win32gui.GetForegroundWindow()
    if not fg_hwnd or not win32gui.IsWindow(fg_hwnd):
        return None

    title = win32gui.GetWindowText(fg_hwnd) or ""
    pid = get_window_pid(fg_hwnd)
    exe = get_process_name(pid)
    if WINDOW_TITLE_KEYWORD.lower() in title.lower() or exe in TARGET_PROCESS_NAMES:
        print(f"[INFO] 使用前台窗口兜底成功 hwnd={fg_hwnd} title={title} exe={exe} pid={pid}")
        return fg_hwnd
    return None


def get_client_origin(hwnd: int) -> Optional[Tuple[int, int]]:
    try:
        origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
        return origin_x, origin_y
    except win32gui.error:
        return None


def build_white_mask(bgr: np.ndarray) -> np.ndarray:
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    lower = np.array([0, 0, 165], dtype=np.uint8)
    upper = np.array([179, 70, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)
    return mask


def match_template_white_mask(roi_bgr: np.ndarray, template_bgr: np.ndarray, template_mask: np.ndarray) -> float:
    """
    多尺度模板匹配（alpha/mask 约束）：
    使用 CCORR_NORMED 避免白掩码低方差导致的 nan/inf。
    返回最大匹配分数（越大越好）。
    """
    roi_gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    best_score = 0.0

    for scale in TEMPLATE_SCALES:
        tw = max(8, int(round(template_bgr.shape[1] * scale)))
        th = max(8, int(round(template_bgr.shape[0] * scale)))
        if roi_gray.shape[1] < tw or roi_gray.shape[0] < th:
            continue

        tpl_resized = cv2.resize(
            template_bgr,
            (tw, th),
            interpolation=cv2.INTER_AREA if scale < 1.0 else cv2.INTER_LINEAR,
        )
        mask_resized = cv2.resize(template_mask, (tw, th), interpolation=cv2.INTER_NEAREST)
        if int(np.count_nonzero(mask_resized)) < 12:
            continue

        tpl_gray = cv2.cvtColor(tpl_resized, cv2.COLOR_BGR2GRAY)
        res = cv2.matchTemplate(roi_gray, tpl_gray, cv2.TM_CCORR_NORMED, mask=mask_resized)
        _min_val, max_val, _min_loc, _max_loc = cv2.minMaxLoc(res)
        if not np.isfinite(max_val):
            continue
        score = float(max_val)
        if score > best_score:
            best_score = score

    return max(0.0, min(1.0, best_score))


def match_template_click(roi_bgr: np.ndarray, template_bgr: np.ndarray, template_mask: np.ndarray) -> float:
    """
    点击图标专用匹配：
    使用 CCOEFF_NORMED，降低纯水面/平坦纹理的高分误命中。
    """
    roi_gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    best_score = -1.0

    for scale in TEMPLATE_SCALES:
        tw = max(8, int(round(template_bgr.shape[1] * scale)))
        th = max(8, int(round(template_bgr.shape[0] * scale)))
        if roi_gray.shape[1] < tw or roi_gray.shape[0] < th:
            continue

        tpl_resized = cv2.resize(
            template_bgr,
            (tw, th),
            interpolation=cv2.INTER_AREA if scale < 1.0 else cv2.INTER_LINEAR,
        )
        mask_resized = cv2.resize(template_mask, (tw, th), interpolation=cv2.INTER_NEAREST)
        if int(np.count_nonzero(mask_resized)) < 12:
            continue

        tpl_gray = cv2.cvtColor(tpl_resized, cv2.COLOR_BGR2GRAY)
        res = cv2.matchTemplate(roi_gray, tpl_gray, cv2.TM_CCOEFF_NORMED, mask=mask_resized)
        _min_val, max_val, _min_loc, _max_loc = cv2.minMaxLoc(res)
        if not np.isfinite(max_val):
            continue
        score = float(max_val)
        if score > best_score:
            best_score = score

    return max(-1.0, min(1.0, best_score))


def match_template_plain(roi_bgr: np.ndarray, template_bgr: np.ndarray) -> float:
    """
    普通模板匹配（不使用任何掩码），用于关闭.png 命中。
    """
    roi_gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    best_score = -1.0

    for scale in TEMPLATE_SCALES:
        tw = max(8, int(round(template_bgr.shape[1] * scale)))
        th = max(8, int(round(template_bgr.shape[0] * scale)))
        if roi_gray.shape[1] < tw or roi_gray.shape[0] < th:
            continue

        tpl_resized = cv2.resize(
            template_bgr,
            (tw, th),
            interpolation=cv2.INTER_AREA if scale < 1.0 else cv2.INTER_LINEAR,
        )
        tpl_gray = cv2.cvtColor(tpl_resized, cv2.COLOR_BGR2GRAY)
        res = cv2.matchTemplate(roi_gray, tpl_gray, cv2.TM_CCORR_NORMED)
        _min_val, max_val, _min_loc, _max_loc = cv2.minMaxLoc(res)
        if not np.isfinite(max_val):
            continue
        score = float(max_val)
        if score > best_score:
            best_score = score

    return max(-1.0, min(1.0, best_score))


def capture_roi_bgr(sct: mss.MSS, client_origin: Tuple[int, int], roi: Tuple[int, int, int, int]) -> np.ndarray:
    ox, oy = client_origin
    left, top, right, bottom = roi
    monitor = {
        "left": ox + left,
        "top": oy + top,
        "width": right - left,
        "height": bottom - top,
    }
    shot = sct.grab(monitor)
    arr = np.array(shot, dtype=np.uint8)
    return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)


def capture_roi_bgr_by_printwindow(hwnd: int, roi: Tuple[int, int, int, int]) -> Optional[np.ndarray]:
    try:
        _left, _top, client_w, client_h = win32gui.GetClientRect(hwnd)
    except win32gui.error:
        return None
    if client_w <= 0 or client_h <= 0:
        return None

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    if hwnd_dc == 0:
        return None

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

                left, top, right, bottom = roi
                left = max(0, left)
                top = max(0, top)
                right = min(client_bgr.shape[1], right)
                bottom = min(client_bgr.shape[0], bottom)
                if right <= left or bottom <= top:
                    continue

                roi_img = client_bgr[top:bottom, left:right]
                if float(np.mean(roi_img)) < 3.0 and float(np.std(roi_img)) < 3.0:
                    continue
                return roi_img
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


def capture_roi_auto(
    sct: mss.MSS,
    hwnd: int,
    client_origin: Tuple[int, int],
    roi: Tuple[int, int, int, int],
    warned_state: dict[str, bool],
) -> Optional[np.ndarray]:
    if CAPTURE_BACKEND == "printwindow":
        img = capture_roi_bgr_by_printwindow(hwnd, roi)
        if img is not None:
            return img
        if ALLOW_MSS_FALLBACK_WHEN_PW_FAIL:
            if not warned_state.get("pw_fallback", False):
                print("[WARN] PrintWindow 抓图失败，回退 mss。")
                warned_state["pw_fallback"] = True
            return capture_roi_bgr(sct, client_origin, roi)
        if not warned_state.get("pw_fail", False):
            print("[WARN] PrintWindow 抓图失败，且已禁用 mss 回退。")
            warned_state["pw_fail"] = True
        return None
    return capture_roi_bgr(sct, client_origin, roi)


def press_f_once() -> None:
    raise RuntimeError("press_f_once 需要传入 hwnd，请使用 press_f_once_to_window(hwnd)")


def fake_activate_window(hwnd: int) -> None:
    win32gui.SendMessage(hwnd, WM_ACTIVATE, WA_ACTIVE, 0)


def send_f_postmessage_simple(hwnd: int, hold_sec: float) -> None:
    """
    与 test_send_f_minimal.py 一致：
    - 先发送伪激活
    - 再最简 PostMessage 发送 F（lParam 固定 0）
    """
    fake_activate_window(hwnd)
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, 0x46, 0)
    time.sleep(hold_sec)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, 0x46, 0)


def press_f_once_to_window(hwnd: int) -> None:
    """按 test_send_f_minimal.py 的策略发送一次 F。"""
    send_f_postmessage_simple(hwnd, F_KEYDOWN_HOLD_SEC)


def send_esc_postmessage_simple(hwnd: int, hold_sec: float) -> None:
    """
    与 F 的发送方式保持一致：
    - 先发送伪激活
    - 再最简 PostMessage 发送 ESC（lParam 固定 0）
    """
    fake_activate_window(hwnd)
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
    time.sleep(hold_sec)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)


def press_esc_once_to_window(hwnd: int) -> None:
    send_esc_postmessage_simple(hwnd, F_KEYDOWN_HOLD_SEC)


def run_flow() -> None:
    click_data = load_template_with_mask(TEMPLATE_CLICK)
    close_tpl = load_template_plain(TEMPLATE_CLOSE)
    if click_data is None or close_tpl is None:
        print("[ERROR] 模板读取失败，请检查 点击/关闭 图片路径。")
        return
    click_tpl, click_mask = click_data

    hwnd = get_window_handle()
    if hwnd is None:
        print("[ERROR] 未找到异环窗口，流程退出。")
        return
    title = win32gui.GetWindowText(hwnd) or ""
    pid = get_window_pid(hwnd)
    exe = get_process_name(pid)

    print("[INFO] 流程启动：等待点击图标 -> 连续F到点击消失 -> 调用fishing_bot -> 等关闭并结束")
    print(f"[INFO] 目标窗口 hwnd={hwnd} pid={pid} exe={exe} title={title}")
    print(f"[INFO] CAPTURE_BACKEND={CAPTURE_BACKEND}, ALLOW_MSS_FALLBACK_WHEN_PW_FAIL={ALLOW_MSS_FALLBACK_WHEN_PW_FAIL}")
    print(f"[INFO] INPUT_BACKEND={INPUT_BACKEND}")

    recovery_state = RecoveryState()
    while True:
        bot_proc: Optional[subprocess.Popen] = None
        bot_start_signal_path = Path(f".runtime/fishing_bot_start_{int(time.time() * 1000)}.signal")
        warned_state: dict[str, bool] = {}
        should_restart_flow = False
        stage_progress_ts = time.time()
        bot_call_start_ts: Optional[float] = None

        def try_recovery(trigger_label: str) -> bool:
            nonlocal should_restart_flow, bot_call_start_ts, stage_progress_ts, bot_proc
            step_stuck_sec = time.time() - stage_progress_ts
            bot_elapsed_sec = 0.0
            if bot_call_start_ts is not None:
                bot_elapsed_sec = time.time() - bot_call_start_ts

            # 需求：当 fishing_bot 调用超过 40s 进入兜底时，先停 bot 再走重启路径。
            if bot_elapsed_sec > BOT_TIMEOUT_SEC and bot_proc is not None and bot_proc.poll() is None:
                print(
                    f"[WARN] fishing_bot 已运行 {bot_elapsed_sec:.2f}s (> {BOT_TIMEOUT_SEC:.2f}s)，"
                    "先停止 fishing_bot，再执行兜底重启。"
                )
                terminate_bot_process(bot_proc)
                bot_proc = None

            def restart_callback() -> bool:
                nonlocal should_restart_flow, bot_proc
                terminate_bot_process(bot_proc)
                bot_proc = None
                should_restart_flow = True
                return True

            recovered = attempt_fishing_recovery(
                state=recovery_state,
                trigger_label=trigger_label,
                step_stuck_sec=step_stuck_sec,
                fishing_bot_elapsed_sec=bot_elapsed_sec,
                restart_callback=restart_callback,
            )
            return recovered

        with mss.MSS() as sct:
            # 1) 点击图标出现
            while True:
                if not win32gui.IsWindow(hwnd):
                    print("[ERROR] 目标窗口失效，流程结束。")
                    return
                recovery_triggered = try_recovery("stage1_wait_click")
                if recovery_triggered:
                    break
                client_origin = get_client_origin(hwnd)
                if client_origin is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                click_roi_img = capture_roi_auto(sct, hwnd, client_origin, CLICK_ROI, warned_state)
                if click_roi_img is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                score_click = match_template_click(click_roi_img, click_tpl, click_mask)
                if score_click >= CLICK_MATCH_THRESHOLD:
                    print(f"[INFO] 命中 点击.png，score={score_click:.3f}，开始按F。")
                    stage_progress_ts = time.time()
                    break
                time.sleep(POLL_INTERVAL)
            if should_restart_flow:
                continue

            # 2) 按F直到“点击图标”无法匹配
            if PRELAUNCH_BOT and bot_proc is None:
                if not PYTHON_EXE.exists() or not FISHING_BOT_PATH.exists():
                    print("[ERROR] .venv Python 或 fishing_bot.py 不存在，无法预启动 fishing_bot.py。")
                    return
                bot_start_signal_path.parent.mkdir(parents=True, exist_ok=True)
                env = os.environ.copy()
                env["FISHING_BOT_WAIT_SIGNAL"] = "1"
                env["FISHING_BOT_START_SIGNAL_PATH"] = str(bot_start_signal_path.resolve())
                bot_proc = subprocess.Popen([str(PYTHON_EXE), str(FISHING_BOT_PATH)], env=env)
                print(f"[INFO] fishing_bot.py 预启动完成，pid={bot_proc.pid}，等待启动信号。")
                stage_progress_ts = time.time()

            absent_count = 0
            f_press_count = 0
            while True:
                recovery_triggered = try_recovery("stage2_press_f")
                if recovery_triggered:
                    break
                client_origin = get_client_origin(hwnd)
                if client_origin is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                click_roi_img = capture_roi_auto(sct, hwnd, client_origin, CLICK_ROI, warned_state)
                if click_roi_img is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                score_click = match_template_click(click_roi_img, click_tpl, click_mask)
                if score_click < CLICK_MATCH_THRESHOLD:
                    absent_count += 1
                else:
                    absent_count = 0
                print(
                    f"[DEBUG] stage2 click_disappear_check "
                    f"score_click={score_click:.3f} "
                    f"threshold={CLICK_MATCH_THRESHOLD:.3f} "
                    f"absent_count={absent_count}/{DISAPPEAR_ABSENT_CONSEC_FRAMES}"
                )

                if absent_count >= DISAPPEAR_ABSENT_CONSEC_FRAMES:
                    print(
                        f"[INFO] 点击图标连续{DISAPPEAR_ABSENT_CONSEC_FRAMES}帧低于阈值，"
                        f"score={score_click:.3f}，开始启动 fishing_bot.py。"
                    )
                    stage_progress_ts = time.time()
                    break

                press_f_once_to_window(hwnd)
                stage_progress_ts = time.time()
                f_press_count += 1
                if f_press_count % 8 == 0:
                    print(
                        f"[DEBUG] stage2 score_click={score_click:.3f} "
                        f"absent_count={absent_count}/{DISAPPEAR_ABSENT_CONSEC_FRAMES} "
                        f"f_press_count={f_press_count}"
                    )
                elif f_press_count <= 3:
                    print(f"[DEBUG] stage2 send_f backend={INPUT_BACKEND} count={f_press_count}")
                if f_press_count >= MAX_F_PRESS_COUNT_IN_STAGE2:
                    print(
                        f"[WARN] stage2 按F已达上限 {MAX_F_PRESS_COUNT_IN_STAGE2} 次，"
                        "强制进入下一阶段并启动 fishing_bot.py，避免循环卡死。"
                    )
                    stage_progress_ts = time.time()
                    break
                time.sleep(F_PRESS_INTERVAL)
            if should_restart_flow:
                continue

            # 3) 调用 fishing_bot.py（预启动模式下发送启动信号）
            if bot_proc is not None and bot_proc.poll() is None and PRELAUNCH_BOT:
                bot_start_signal_path.write_text("start\n", encoding="utf-8")
                print(f"[INFO] 已发送 fishing_bot.py 启动信号：{bot_start_signal_path}")
            else:
                if not PYTHON_EXE.exists() or not FISHING_BOT_PATH.exists():
                    print("[ERROR] .venv Python 或 fishing_bot.py 不存在，无法调用。")
                    return
                bot_proc = subprocess.Popen([str(PYTHON_EXE), str(FISHING_BOT_PATH)])
                print(f"[INFO] fishing_bot.py 已启动，pid={bot_proc.pid}")
            bot_call_start_ts = time.time()
            stage_progress_ts = bot_call_start_ts

            # 4) 结束标志：匹配关闭.png 或 点击.png（点击仍在 CLICK_ROI 轮询）
            stage4_click_hit_count = 0
            stage4_close_hit_count = 0
            while True:
                recovery_triggered = try_recovery("stage4_wait_end")
                if recovery_triggered:
                    break
                if bot_proc.poll() is not None:
                    print("[WARN] fishing_bot.py 已提前退出。")
                    stage_progress_ts = time.time()
                    break
                client_origin = get_client_origin(hwnd)
                if client_origin is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                close_roi_img = capture_roi_auto(sct, hwnd, client_origin, CLOSE_ROI, warned_state)
                if close_roi_img is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                score_close = match_template_plain(close_roi_img, close_tpl)
                if score_close >= CLOSE_MATCH_THRESHOLD:
                    stage4_close_hit_count += 1
                else:
                    stage4_close_hit_count = 0
                if stage4_close_hit_count >= CLOSE_HIT_CONSEC_FRAMES:
                    print(
                        f"[INFO] 连续{CLOSE_HIT_CONSEC_FRAMES}帧命中 关闭.png，"
                        f"latest_score={score_close:.3f}，准备结束 fishing_bot.py"
                    )
                    terminate_bot_process(bot_proc)
                    press_esc_once_to_window(hwnd)
                    print("[INFO] 已发送一次 ESC。")
                    stage_progress_ts = time.time()
                    break

                click_roi_img = capture_roi_auto(sct, hwnd, client_origin, CLICK_ROI, warned_state)
                if click_roi_img is None:
                    time.sleep(POLL_INTERVAL)
                    continue
                score_click = match_template_click(click_roi_img, click_tpl, click_mask)
                if score_click >= CLICK_MATCH_THRESHOLD:
                    stage4_click_hit_count += 1
                else:
                    stage4_click_hit_count = 0

                if stage4_click_hit_count >= 3:
                    print(
                        f"[INFO] 连续3帧命中 点击.png，latest_score={score_click:.3f}，准备结束 fishing_bot.py"
                    )
                    terminate_bot_process(bot_proc)
                    # 命中点击后不立即按 ESC：等待 3 秒，再看是否出现关闭按钮
                    time.sleep(3.0)
                    close_hit_count_after_wait = 0
                    last_score_close_after_wait = -1.0
                    close_confirmed_after_wait = False
                    for _ in range(CLOSE_HIT_CONSEC_FRAMES + 2):
                        client_origin_after_wait = get_client_origin(hwnd)
                        if client_origin_after_wait is None:
                            close_hit_count_after_wait = 0
                            time.sleep(POLL_INTERVAL)
                            continue
                        close_roi_after_wait = capture_roi_auto(
                            sct, hwnd, client_origin_after_wait, CLOSE_ROI, warned_state
                        )
                        if close_roi_after_wait is None:
                            close_hit_count_after_wait = 0
                            time.sleep(POLL_INTERVAL)
                            continue

                        score_close_after_wait = match_template_plain(close_roi_after_wait, close_tpl)
                        last_score_close_after_wait = score_close_after_wait
                        if score_close_after_wait >= CLOSE_MATCH_THRESHOLD:
                            close_hit_count_after_wait += 1
                        else:
                            close_hit_count_after_wait = 0

                        if close_hit_count_after_wait >= CLOSE_HIT_CONSEC_FRAMES:
                            close_confirmed_after_wait = True
                            break
                        time.sleep(POLL_INTERVAL)

                    if close_confirmed_after_wait:
                        press_esc_once_to_window(hwnd)
                        print(
                            f"[INFO] 点击后等待3s连续{CLOSE_HIT_CONSEC_FRAMES}帧命中关闭.png，"
                            f"latest_score={last_score_close_after_wait:.3f}，已发送一次 ESC。"
                        )
                    else:
                        if last_score_close_after_wait >= 0.0:
                            print(
                                f"[INFO] 点击后等待3s未连续{CLOSE_HIT_CONSEC_FRAMES}帧命中关闭.png，"
                                f"latest_score={last_score_close_after_wait:.3f}，不发送 ESC。"
                            )
                        else:
                            print(
                                f"[INFO] 点击后等待3s未连续{CLOSE_HIT_CONSEC_FRAMES}帧命中关闭.png，"
                                "且关闭区域截图无效，不发送 ESC。"
                            )
                    stage_progress_ts = time.time()
                    break
                time.sleep(POLL_INTERVAL)
            if should_restart_flow:
                print("[WARN] 已触发兜底恢复，重新拉起主流程。")
                continue
            break

    print("[INFO] 新流程结束。")


if __name__ == "__main__":
    run_flow()
