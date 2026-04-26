#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import threading
import time
from pathlib import Path
from typing import List, Optional, Tuple

import cv2
import win32api
import win32con
import win32gui

import make_coffee_by_image as coffee
import 大锤模式 as hammer_mode


MATCH_THRESHOLD = 0.84
POLL_INTERVAL_SEC = 0.12
CLICK_INTERVAL_SEC = 0.12
WAIT_AFTER_START_SEC = 50.0
START_DISAPPEAR_TIMEOUT_SEC = 20.0
POST_FIRST_CLICK_PRIORITY_SEC = 3.0
ASSETS_DIR = coffee.BASE_DIR / "素材"
MANAGER_SPECIAL_PATH = ASSETS_DIR / "店长特供.png"
START_TEMPLATE_PATH = ASSETS_DIR / "开始营业.png"
CLAIM_TEMPLATE_PATH = ASSETS_DIR / "领取.png"
EXIT_TEMPLATE_PATH = ASSETS_DIR / "退出.png"


def send_f_key(hwnd: int) -> None:
    """
    向目标窗口发送一次 F 按键（按下+抬起）。
    """
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.05)
    except Exception:
        pass
    vk_f = 0x46
    win32api.keybd_event(vk_f, 0, 0, 0)
    time.sleep(0.03)
    win32api.keybd_event(vk_f, 0, win32con.KEYEVENTF_KEYUP, 0)


def load_scaled_templates(path: Path) -> List:
    base = coffee.imread_unicode(path)
    if base is None:
        raise ValueError(f"模板读取失败: {path}")
    return coffee.build_scaled_templates([base], (0.90, 0.95, 1.00, 1.05, 1.10))


def screenshot_gray(hwnd: int) -> Optional:
    frame = coffee.screenshot_client_bgr(hwnd)
    if frame is None:
        return None
    return cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)


def detect_template_center(
    hwnd: int,
    templates: List,
) -> Tuple[float, Optional[Tuple[int, int]]]:
    gray = screenshot_gray(hwnd)
    if gray is None:
        return 0.0, None
    return coffee.max_match_score_with_center(gray, templates)


def wait_and_click_template(hwnd: int, templates: List, name: str, timeout_sec: Optional[float] = None) -> None:
    start_ts = time.time()
    while True:
        score, center = detect_template_center(hwnd, templates)
        if score >= MATCH_THRESHOLD and center is not None:
            print(f"命中模板[{name}] score={score:.3f}，点击中心点 {center}")
            coffee.click_rel(hwnd, center[0], center[1])
            return

        if timeout_sec is not None and (time.time() - start_ts) >= timeout_sec:
            raise TimeoutError(f"等待模板[{name}]超时({timeout_sec}s)，当前最高score={score:.3f}")
        time.sleep(POLL_INTERVAL_SEC)


def wait_manager_then_press_f_until_start(
    hwnd: int,
    manager_templates: List,
    start_templates: List,
) -> None:
    print("步骤1: 先轮询 店长特供")
    while True:
        manager_score, _ = detect_template_center(hwnd, manager_templates)
        if manager_score >= MATCH_THRESHOLD:
            print(f"命中模板[店长特供] score={manager_score:.3f}，开始按 F 触发开始营业。")
            break
        time.sleep(POLL_INTERVAL_SEC)

    print("步骤2: 连续按 F，直到出现 开始营业 并点击")
    while True:
        start_score, start_center = detect_template_center(hwnd, start_templates)
        if start_score >= MATCH_THRESHOLD and start_center is not None:
            print(f"命中模板[开始营业] score={start_score:.3f}，点击中心点 {start_center}")
            coffee.click_rel(hwnd, start_center[0], start_center[1])
            return
        send_f_key(hwnd)
        time.sleep(0.20)


def run_coffee_worker() -> None:
    # 直接复用主脚本循环，和营业流程并行执行。
    coffee.main()


def run_hammer_worker() -> None:
    # 复用大锤脚本主循环，作为并行流程。
    raise RuntimeError("run_hammer_worker 需要通过带 run_gate 的版本调用。")


def run_hammer_worker_with_gate(run_gate: threading.Event) -> None:
    hammer_mode.main(auto_start=True, run_gate=run_gate)


def wait_start_template_disappear(
    hwnd: int,
    start_templates: List,
    timeout_sec: float = START_DISAPPEAR_TIMEOUT_SEC,
) -> None:
    start_ts = time.time()
    while True:
        start_score, _ = detect_template_center(hwnd, start_templates)
        if start_score < MATCH_THRESHOLD:
            print(f"步骤3: 检测到开始营业已消失（score={start_score:.3f}），放行大锤。")
            return
        if (time.time() - start_ts) >= timeout_sec:
            print(
                f"步骤3: 等待开始营业消失超时({timeout_sec:.0f}s, score={start_score:.3f})，"
                "继续放行大锤。"
            )
            return
        time.sleep(POLL_INTERVAL_SEC)


def run_single_round(
    manager_templates: List,
    start_templates: List,
    claim_templates: List,
    exit_templates: List,
    round_index: int,
    wait_after_start_sec: float,
    worker_mode: str,
    hammer_run_gate: Optional[threading.Event] = None,
) -> None:
    print(f"\n===== 第 {round_index} 轮营业流程开始 =====")
    hwnd = coffee.pick_game_window()
    if hwnd is None:
        raise RuntimeError("未找到异环窗口。")

    if hammer_run_gate is not None:
        hammer_run_gate.clear()
    wait_manager_then_press_f_until_start(hwnd, manager_templates, start_templates)

    if worker_mode == "hammer" and hammer_run_gate is not None:
        wait_start_template_disappear(hwnd, start_templates)
        print(f"步骤4: 开始营业消失后，等待 {wait_after_start_sec:.0f} 秒（此段放行大锤）")
        hammer_run_gate.set()
        try:
            time.sleep(wait_after_start_sec)
        finally:
            hammer_run_gate.clear()
        print("步骤5: 连续点击(40,40)直到出现 领取或退出，再点击")
    else:
        print(f"步骤3: 点击开始营业后等待 {wait_after_start_sec:.0f} 秒")
        time.sleep(wait_after_start_sec)
        print("步骤4: 连续点击(40,40)直到出现 领取或退出，再点击")

    first_click_ts: Optional[float] = None
    while True:
        claim_score, claim_center = detect_template_center(hwnd, claim_templates)
        if claim_score >= MATCH_THRESHOLD and claim_center is not None:
            print(f"命中 领取 score={claim_score:.3f}，点击 {claim_center}")
            coffee.click_rel(hwnd, claim_center[0], claim_center[1])
            print(f"第 {round_index} 轮完成。")
            return

        exit_score, exit_center = detect_template_center(hwnd, exit_templates)
        if exit_score >= MATCH_THRESHOLD and exit_center is not None:
            print(f"命中 退出 score={exit_score:.3f}，点击 {exit_center}")
            coffee.click_rel(hwnd, exit_center[0], exit_center[1])
            print(f"第 {round_index} 轮完成（退出）。")
            return

        if first_click_ts is None:
            coffee.click_rel(hwnd, 40, 40)
            first_click_ts = time.time()
            print(f"首次点击(40,40)后，进入 {POST_FIRST_CLICK_PRIORITY_SEC:.0f}s 领取优先轮询窗口。")
            time.sleep(POLL_INTERVAL_SEC)
            continue

        if (time.time() - first_click_ts) <= POST_FIRST_CLICK_PRIORITY_SEC:
            # 首次点击后的短窗口内，不再追加点击，优先高频检测领取，其次检测退出。
            time.sleep(POLL_INTERVAL_SEC)
            continue

        coffee.click_rel(hwnd, 40, 40)
        time.sleep(CLICK_INTERVAL_SEC)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="异环营业流程脚本")
    parser.add_argument(
        "--worker-mode",
        choices=("coffee", "hammer"),
        default="coffee",
        help="并行工作线程模式：coffee=make_coffee_by_image，hammer=大锤模式",
    )
    parser.add_argument(
        "--wait-after-start-sec",
        type=float,
        default=WAIT_AFTER_START_SEC,
        help="点击开始营业后的等待秒数",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    dpi_mode = coffee.enable_dpi_awareness()
    print(f"DPI 感知模式: {dpi_mode}")
    print("提醒：当前流程按 1600x900 固定坐标运行（未做16:9归一化）。")
    print("提醒：本流程将持续循环，按 F10 可停止。")
    print(f"当前并行模式: {args.worker_mode}")
    if args.wait_after_start_sec < 0:
        raise ValueError("--wait-after-start-sec 不能小于 0")
    print(f"开始营业后等待时长: {args.wait_after_start_sec:.1f}s")

    for p in (MANAGER_SPECIAL_PATH, START_TEMPLATE_PATH, CLAIM_TEMPLATE_PATH, EXIT_TEMPLATE_PATH):
        if not p.exists():
            raise FileNotFoundError(f"模板不存在: {p}")

    manager_templates = load_scaled_templates(MANAGER_SPECIAL_PATH)
    start_templates = load_scaled_templates(START_TEMPLATE_PATH)
    claim_templates = load_scaled_templates(CLAIM_TEMPLATE_PATH)
    exit_templates = load_scaled_templates(EXIT_TEMPLATE_PATH)
    hammer_run_gate: Optional[threading.Event] = None

    if args.worker_mode == "hammer":
        print("启动 大锤模式 并行流程（仅启动一次）")
        hammer_run_gate = threading.Event()
        hammer_run_gate.clear()
        worker = threading.Thread(target=run_hammer_worker_with_gate, args=(hammer_run_gate,), daemon=True)
    else:
        print("启动 make_coffee_by_image 并行流程（仅启动一次）")
        worker = threading.Thread(target=run_coffee_worker, daemon=True)
    worker.start()
    time.sleep(0.30)

    round_index = 1
    while True:
        run_single_round(
            manager_templates=manager_templates,
            start_templates=start_templates,
            claim_templates=claim_templates,
            exit_templates=exit_templates,
            round_index=round_index,
            wait_after_start_sec=args.wait_after_start_sec,
            worker_mode=args.worker_mode,
            hammer_run_gate=hammer_run_gate,
        )
        round_index += 1
        time.sleep(0.5)


if __name__ == "__main__":
    main()
