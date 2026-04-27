#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import time
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple

import cv2
import mss
import numpy as np
import pywintypes
import win32api
import win32con
import win32gui
import win32process
import psutil


# 固定 1600x900 客户区基准，不做归一化
WINDOW_TITLE_KEYWORD = "异环"
TARGET_PROCESS_NAMES = {"HTGame.exe"}
TEMPLATE_START_FISHING = Path("素材/钓鱼/开始钓鱼.png")
TEMPLATE_CLOSE = Path("素材/钓鱼/关闭.png")
TEMPLATE_CLICK = Path("素材/钓鱼/点击.png")
TEMPLATE_CONFIRM = Path("素材/钓鱼/确认.png")

POLL_INTERVAL = 0.10
ESC_HOLD_SEC = 0.08
MATCH_THRESHOLD = 0.82
BLOCK_TIMEOUT_SEC = 10.0
BOT_TIMEOUT_SEC = 40.0
COUNTER_RESET_WINDOW_SEC = 60.0
MAX_CONSECUTIVE_RECOVERY = 3
RECOVERY_RETRY_COOLDOWN_SEC = 2.0
RESTART_CALLBACK_TIMEOUT_SEC = 5.0
WM_ACTIVATE = 0x0006
WA_ACTIVE = 1
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
MK_LBUTTON = 0x0001
WM_NCHITTEST = 0x0084
WM_SETCURSOR = 0x0020
HTCLIENT = 0x0001
START_FISHING_CLICK_X = 800
START_FISHING_CLICK_Y = 400
FOREGROUND_CLICK_HOLD_SEC = 0.08


@dataclass
class RecoveryState:
    consecutive_recovery_count: int = 0
    last_restart_ts: Optional[float] = None
    next_recovery_check_ts: float = 0.0


def imread_unicode(path: Path) -> Optional[np.ndarray]:
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except Exception:
        return None


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
    except win32gui.error:
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
        return fg_hwnd
    return None


def get_client_origin(hwnd: int) -> Optional[Tuple[int, int]]:
    try:
        return win32gui.ClientToScreen(hwnd, (0, 0))
    except win32gui.error:
        return None


def fake_activate_window(hwnd: int) -> None:
    win32gui.SendMessage(hwnd, WM_ACTIVATE, WA_ACTIVE, 0)


def send_esc_postmessage_simple(hwnd: int, hold_sec: float) -> None:
    fake_activate_window(hwnd)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
        time.sleep(hold_sec)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)
    except pywintypes.error as exc:
        if len(exc.args) >= 1 and exc.args[0] == 5:
            raise RuntimeError(
                "后台键入ESC被系统拒绝访问（UIPI）。请用管理员权限启动当前Python/IDE，"
                "并确保与游戏进程权限级别一致；否则只能改为前台输入方案。"
            ) from exc
        raise


def click_rel_postmessage(hwnd: int, rel_x: int, rel_y: int) -> None:
    fake_activate_window(hwnd)
    lparam = (int(rel_y) << 16) | (int(rel_x) & 0xFFFF)
    try:
        win32api.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
        time.sleep(0.01)
        win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
        time.sleep(0.01)
        win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lparam)
    except pywintypes.error as exc:
        if len(exc.args) >= 1 and exc.args[0] == 5:
            raise RuntimeError(
                "后台点击被系统拒绝访问（UIPI）。请用管理员权限启动当前Python/IDE，"
                "并确保与游戏进程权限级别一致；否则只能改为前台点击方案。"
            ) from exc
        raise


def background_hover(hwnd: int, x: int, y: int) -> None:
    """后台悬停：仅发送鼠标移动消息，不移动真实鼠标。"""
    fake_activate_window(hwnd)
    lparam = (int(y) << 16) | (int(x) & 0xFFFF)  # MAKELONG(x, y)
    try:
        win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
    except pywintypes.error as exc:
        if len(exc.args) >= 1 and exc.args[0] == 5:
            raise RuntimeError(
                "后台悬停被系统拒绝访问（UIPI）。请用管理员权限启动当前Python/IDE，"
                "并确保与游戏进程权限级别一致。"
            ) from exc
        raise


def background_left_click(hwnd: int, x: int, y: int, hold_sec: float = 1.0) -> None:
    """
    最小后台左键点击：
    - 不移动真实鼠标，不切前台
    - 使用客户端坐标
    - 使用 win32gui.PostMessage 发送按下/抬起
    """
    fake_activate_window(hwnd)
    cx, cy = int(x), int(y)
    lparam_client = (cy << 16) | (cx & 0xFFFF)  # MAKELONG(x, y) 客户区坐标
    try:
        # 1) 鼠标移动到目标点（客户区）
        win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lparam_client)
        time.sleep(0.01)

        # 2) 命中测试（屏幕坐标）
        sx, sy = win32gui.ClientToScreen(hwnd, (cx, cy))
        lparam_screen = (int(sy) << 16) | (int(sx) & 0xFFFF)
        hit_test = win32gui.SendMessage(hwnd, WM_NCHITTEST, 0, lparam_screen)
        if hit_test == 0:
            hit_test = HTCLIENT

        # 3) 设置光标上下文（把命中结果 + 当前消息类型传回窗口）
        setcursor_lparam = ((win32con.WM_MOUSEMOVE & 0xFFFF) << 16) | (hit_test & 0xFFFF)
        win32gui.SendMessage(hwnd, WM_SETCURSOR, hwnd, setcursor_lparam)
        time.sleep(0.01)

        # 4) 再执行按下/抬起
        win32gui.PostMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lparam_client)
        time.sleep(hold_sec)
        win32gui.PostMessage(hwnd, WM_LBUTTONUP, 0, lparam_client)
    except pywintypes.error as exc:
        if len(exc.args) >= 1 and exc.args[0] == 5:
            raise RuntimeError(
                "后台点击被系统拒绝访问（UIPI）。请用管理员权限启动当前Python/IDE，"
                "并确保与游戏进程权限级别一致；否则只能改为前台点击方案。"
            ) from exc
        raise


def capture_client_bgr(sct: mss.MSS, hwnd: int) -> Optional[np.ndarray]:
    origin = get_client_origin(hwnd)
    if origin is None:
        return None
    left, top = origin
    try:
        _l, _t, width, height = win32gui.GetClientRect(hwnd)
    except win32gui.error:
        return None
    if width <= 0 or height <= 0:
        return None
    monitor = {"left": left, "top": top, "width": width, "height": height}
    shot = sct.grab(monitor)
    arr = np.array(shot, dtype=np.uint8)
    return cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)


def find_template_center(
    frame_bgr: np.ndarray,
    template_bgr: np.ndarray,
    threshold: float,
) -> Optional[Tuple[int, int, float]]:
    frame_gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)
    tpl_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
    if frame_gray.shape[0] < tpl_gray.shape[0] or frame_gray.shape[1] < tpl_gray.shape[1]:
        return None
    res = cv2.matchTemplate(frame_gray, tpl_gray, cv2.TM_CCOEFF_NORMED)
    _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(res)
    score = float(max_val)
    if score < threshold:
        return None
    x = max_loc[0] + tpl_gray.shape[1] // 2
    y = max_loc[1] + tpl_gray.shape[0] // 2
    return x, y, score


def calc_template_score(
    frame_bgr: np.ndarray,
    template_bgr: np.ndarray,
) -> float:
    frame_gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)
    tpl_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)
    if frame_gray.shape[0] < tpl_gray.shape[0] or frame_gray.shape[1] < tpl_gray.shape[1]:
        return -1.0
    res = cv2.matchTemplate(frame_gray, tpl_gray, cv2.TM_CCOEFF_NORMED)
    _min_val, max_val, _min_loc, _max_loc = cv2.minMaxLoc(res)
    return float(max_val)


def should_trigger_recovery(
    step_stuck_sec: float,
    fishing_bot_elapsed_sec: float,
) -> bool:
    return step_stuck_sec >= BLOCK_TIMEOUT_SEC or fishing_bot_elapsed_sec > BOT_TIMEOUT_SEC


def _load_templates() -> Dict[str, np.ndarray]:
    start_tpl = imread_unicode(TEMPLATE_START_FISHING)
    close_tpl = imread_unicode(TEMPLATE_CLOSE)
    click_tpl = imread_unicode(TEMPLATE_CLICK)
    confirm_tpl = imread_unicode(TEMPLATE_CONFIRM)
    missing = []
    if start_tpl is None:
        missing.append(str(TEMPLATE_START_FISHING))
    if close_tpl is None:
        missing.append(str(TEMPLATE_CLOSE))
    if click_tpl is None:
        missing.append(str(TEMPLATE_CLICK))
    if confirm_tpl is None:
        missing.append(str(TEMPLATE_CONFIRM))
    if missing:
        raise FileNotFoundError("模板读取失败: " + ", ".join(missing))
    return {"start": start_tpl, "close": close_tpl, "click": click_tpl, "confirm": confirm_tpl}


def _detect_recovery_target(
    frame_bgr: np.ndarray,
    templates: Dict[str, np.ndarray],
    threshold: float = MATCH_THRESHOLD,
) -> Tuple[str, float]:
    # 优先 close，再 click；避免误判时跳过应先退出的弹层。
    close_hit = find_template_center(frame_bgr, templates["close"], threshold)
    if close_hit is not None:
        return "close", close_hit[2]
    click_hit = find_template_center(frame_bgr, templates["click"], threshold)
    if click_hit is not None:
        return "click", click_hit[2]
    return "none", 0.0


def _bring_window_foreground(hwnd: int) -> None:
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # 某些系统会限制前置焦点，失败时仍继续尝试点击。
        pass


def _foreground_click_rel(hwnd: int, rel_x: int, rel_y: int, hold_sec: float = FOREGROUND_CLICK_HOLD_SEC) -> None:
    origin = get_client_origin(hwnd)
    if origin is None:
        raise RuntimeError("获取窗口客户区原点失败，无法前台点击。")
    abs_x = origin[0] + int(rel_x)
    abs_y = origin[1] + int(rel_y)
    _bring_window_foreground(hwnd)
    win32api.SetCursorPos((abs_x, abs_y))
    time.sleep(0.02)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(max(0.01, float(hold_sec)))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def _record_recovery_failure(state: RecoveryState, reason: str) -> None:
    state.consecutive_recovery_count += 1
    state.last_restart_ts = time.time()
    state.next_recovery_check_ts = state.last_restart_ts + RECOVERY_RETRY_COOLDOWN_SEC
    print(
        "[ERROR] 兜底失败："
        f"{reason}。restart_fail_count={state.consecutive_recovery_count}/{MAX_CONSECUTIVE_RECOVERY}"
    )
    if state.consecutive_recovery_count >= MAX_CONSECUTIVE_RECOVERY:
        raise RuntimeError("兜底连续失败达到上限，结束进程。")


def _safe_restart_callback(restart_callback: Callable[[], bool], timeout_sec: float = RESTART_CALLBACK_TIMEOUT_SEC) -> bool:
    result_holder: dict[str, object] = {"done": False, "ok": False, "err": None}

    def _runner() -> None:
        try:
            result_holder["ok"] = bool(restart_callback())
        except Exception as exc:  # pragma: no cover - 防御性分支
            result_holder["err"] = exc
            result_holder["ok"] = False
        finally:
            result_holder["done"] = True

    th = threading.Thread(target=_runner, daemon=True)
    th.start()
    th.join(timeout=max(0.1, float(timeout_sec)))

    if not bool(result_holder["done"]):
        print(f"[ERROR] 兜底失败：重启回调执行超时({timeout_sec:.1f}s)，已中断等待。")
        return False
    if result_holder["err"] is not None:
        print(f"[ERROR] 兜底失败：重启回调抛异常：{result_holder['err']}")
        return False
    return bool(result_holder["ok"])


def _run_foreground_fishing_chain(
    hwnd: int,
    sct: mss.MSS,
    templates: Dict[str, np.ndarray],
    threshold: float,
) -> bool:
    step_a_deadline = time.time() + BLOCK_TIMEOUT_SEC
    while time.time() < step_a_deadline:
        frame_bgr = capture_client_bgr(sct, hwnd)
        if frame_bgr is None:
            time.sleep(POLL_INTERVAL)
            continue
        confirm_hit = find_template_center(frame_bgr, templates["confirm"], threshold)
        if confirm_hit is not None:
            print(f"[INFO] 前台兜底A完成：命中确认.png score={confirm_hit[2]:.3f}")
            break
        start_hit = find_template_center(frame_bgr, templates["start"], threshold)
        if start_hit is not None:
            _foreground_click_rel(hwnd, start_hit[0], start_hit[1])
            print(f"[DEBUG] 前台点击开始钓鱼，score={start_hit[2]:.3f}")
        time.sleep(POLL_INTERVAL)
    else:
        return False

    step_b_deadline = time.time() + BLOCK_TIMEOUT_SEC
    while time.time() < step_b_deadline:
        frame_bgr = capture_client_bgr(sct, hwnd)
        if frame_bgr is None:
            time.sleep(POLL_INTERVAL)
            continue
        click_hit = find_template_center(frame_bgr, templates["click"], threshold)
        if click_hit is not None:
            print(f"[INFO] 前台兜底B完成：命中点击.png score={click_hit[2]:.3f}")
            return True
        confirm_hit = find_template_center(frame_bgr, templates["confirm"], threshold)
        if confirm_hit is not None:
            _foreground_click_rel(hwnd, confirm_hit[0], confirm_hit[1])
            print(f"[DEBUG] 前台点击确认，score={confirm_hit[2]:.3f}")
        time.sleep(POLL_INTERVAL)
    return False


def _refresh_consecutive_window(state: RecoveryState, now_ts: float) -> None:
    if state.last_restart_ts is None:
        return
    if now_ts - state.last_restart_ts > COUNTER_RESET_WINDOW_SEC:
        state.consecutive_recovery_count = 0


def attempt_fishing_recovery(
    state: RecoveryState,
    trigger_label: str,
    step_stuck_sec: float,
    fishing_bot_elapsed_sec: float,
    restart_callback: Callable[[], bool],
    threshold: float = MATCH_THRESHOLD,
) -> bool:
    if not should_trigger_recovery(step_stuck_sec, fishing_bot_elapsed_sec):
        return False

    now_ts = time.time()
    if now_ts < state.next_recovery_check_ts:
        return False
    _refresh_consecutive_window(state, now_ts)

    hwnd = get_window_handle()
    if hwnd is None or (not win32gui.IsWindow(hwnd)):
        print("[ERROR] 兜底失败：未找到有效异环窗口。")
        return False

    templates = _load_templates()
    with mss.MSS() as sct:
        frame_bgr = capture_client_bgr(sct, hwnd)
        if frame_bgr is None:
            print("[ERROR] 兜底失败：截图失败。")
            return False

        matched, score = _detect_recovery_target(frame_bgr, templates, threshold=threshold)
        print(
            "[INFO] recovery_trigger="
            f"{trigger_label}, step_stuck={step_stuck_sec:.2f}s, "
            f"bot_elapsed={fishing_bot_elapsed_sec:.2f}s, matched={matched}, score={score:.3f}, "
            f"count={state.consecutive_recovery_count}"
        )
        if matched == "none":
            start_hit = find_template_center(frame_bgr, templates["start"], threshold)
            if start_hit is None:
                state.next_recovery_check_ts = now_ts + RECOVERY_RETRY_COOLDOWN_SEC
                close_score_now = calc_template_score(frame_bgr, templates["close"])
                click_score_now = calc_template_score(frame_bgr, templates["click"])
                start_score_now = calc_template_score(frame_bgr, templates["start"])
                confirm_score_now = calc_template_score(frame_bgr, templates["confirm"])
                print(
                    "[WARN] 兜底未命中关闭/点击/开始钓鱼模板，本轮恢复失败。"
                    f"score_now(close={close_score_now:.3f}, "
                    f"click={click_score_now:.3f}, "
                    f"start={start_score_now:.3f}, "
                    f"confirm={confirm_score_now:.3f}, "
                    f"threshold={threshold:.3f})"
                )
                return False
            print(f"[INFO] 命中开始钓鱼.png，进入前台兜底链路，score={start_hit[2]:.3f}")
            _bring_window_foreground(hwnd)
            restarted = _run_foreground_fishing_chain(hwnd, sct, templates, threshold)
            if not restarted:
                _record_recovery_failure(state, "前台兜底链路超时(10s+10s)")
                return False
            # 该链路成功后已经回到“点击.png”可继续状态，不再强制重启主流程。
            state.last_restart_ts = time.time()
            state.next_recovery_check_ts = state.last_restart_ts + RECOVERY_RETRY_COOLDOWN_SEC
            state.consecutive_recovery_count = 0
            print("[INFO] 前台兜底完成：已恢复到点击态，继续当前流程。")
            return False

    if matched == "close":
        send_esc_postmessage_simple(hwnd, ESC_HOLD_SEC)
        print("[INFO] 命中关闭.png，已发送ESC一次。")
    else:
        print("[INFO] 命中点击.png，按规则触发重启流程。")

    restarted = _safe_restart_callback(restart_callback)
    if restarted:
        state.last_restart_ts = time.time()
        state.next_recovery_check_ts = 0.0
        state.consecutive_recovery_count = 0
        print("[INFO] 兜底完成：已重新拉起钓鱼流程。")
    else:
        _record_recovery_failure(state, "拉起流程返回失败")
    return restarted


if __name__ == "__main__":
    # 本地联调入口：仅验证恢复触发与模板识别逻辑是否可执行。
    demo_state = RecoveryState()

    def _demo_restart() -> bool:
        print("[DEMO] restart_callback invoked.")
        return True

    ok = attempt_fishing_recovery(
        state=demo_state,
        trigger_label="demo",
        step_stuck_sec=11.0,
        fishing_bot_elapsed_sec=42.0,
        restart_callback=_demo_restart,
    )
    raise SystemExit(0 if ok else 1)
