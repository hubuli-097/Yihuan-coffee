#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
从钓鱼测试图自动测算绿色区域与黄色竖线的 HSV 建议阈值。

默认输入图片：
    素材/钓鱼/钓鱼测试.png

运行：
    .\.venv\Scripts\python.exe calc_fishing_hsv.py
或：
    .\.venv\Scripts\python.exe calc_fishing_hsv.py "素材/钓鱼/钓鱼测试.png"
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Tuple

import cv2
import numpy as np


DEFAULT_IMAGE = Path("素材/钓鱼/钓鱼测试.png")
OUTPUT_DIR = Path("数据记录/调试/calc_fishing_hsv")
OUT_GREEN_MASK = OUTPUT_DIR / "hsv_green_mask.png"
OUT_YELLOW_MASK = OUTPUT_DIR / "hsv_yellow_mask.png"
OUT_OVERLAY = OUTPUT_DIR / "hsv_debug_overlay.png"


def imread_unicode(path: Path) -> np.ndarray | None:
    """兼容中文路径读取图片。"""
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except Exception:
        return None


def percentile_hsv(hsv_pixels: np.ndarray, low_p: float = 5, high_p: float = 95) -> Tuple[np.ndarray, np.ndarray]:
    """按分位数给出 HSV 下上界。"""
    lower = np.percentile(hsv_pixels, low_p, axis=0).astype(np.int32)
    upper = np.percentile(hsv_pixels, high_p, axis=0).astype(np.int32)
    return lower, upper


def clamp_hsv(lower: np.ndarray, upper: np.ndarray, h_pad: int, s_pad: int, v_pad: int) -> Tuple[np.ndarray, np.ndarray]:
    """给分位区间加安全边距，并限制在 OpenCV HSV 范围。"""
    lo = np.array([max(0, lower[0] - h_pad), max(0, lower[1] - s_pad), max(0, lower[2] - v_pad)], dtype=np.uint8)
    hi = np.array([min(179, upper[0] + h_pad), min(255, upper[1] + s_pad), min(255, upper[2] + v_pad)], dtype=np.uint8)
    return lo, hi


def keep_largest_component(mask: np.ndarray) -> np.ndarray:
    """仅保留最大连通域，减少背景干扰。"""
    num_labels, labels, stats, _centroids = cv2.connectedComponentsWithStats(mask, connectivity=8)
    if num_labels <= 1:
        return mask
    largest = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    out = np.zeros_like(mask)
    out[labels == largest] = 255
    return out


def main() -> None:
    image_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_IMAGE
    img = imread_unicode(image_path)
    if img is None:
        print(f"[ERROR] 图片读取失败：{image_path}")
        return

    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # 先用宽松先验范围做初筛，再做统计收敛
    seed_green = cv2.inRange(hsv, np.array([30, 40, 40], np.uint8), np.array([100, 255, 255], np.uint8))
    seed_yellow = cv2.inRange(hsv, np.array([10, 50, 80], np.uint8), np.array([45, 255, 255], np.uint8))

    seed_green = keep_largest_component(seed_green)
    seed_yellow = keep_largest_component(seed_yellow)

    green_pixels = hsv[seed_green > 0]
    yellow_pixels = hsv[seed_yellow > 0]
    if green_pixels.size == 0 or yellow_pixels.size == 0:
        print("[ERROR] 未能从图片中稳定提取绿色/黄色像素，请检查样本图。")
        return

    g_lo_p, g_hi_p = percentile_hsv(green_pixels, 5, 95)
    y_lo_p, y_hi_p = percentile_hsv(yellow_pixels, 5, 95)

    # 绿色给稍大 H 边距，黄色给更紧 H 边距
    g_lo, g_hi = clamp_hsv(g_lo_p, g_hi_p, h_pad=6, s_pad=25, v_pad=25)
    y_lo, y_hi = clamp_hsv(y_lo_p, y_hi_p, h_pad=4, s_pad=20, v_pad=20)

    green_mask = cv2.inRange(hsv, g_lo, g_hi)
    yellow_mask = cv2.inRange(hsv, y_lo, y_hi)

    green_mask = cv2.morphologyEx(green_mask, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8), iterations=1)
    yellow_mask = cv2.morphologyEx(yellow_mask, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8), iterations=1)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    cv2.imwrite(str(OUT_GREEN_MASK), green_mask)
    cv2.imwrite(str(OUT_YELLOW_MASK), yellow_mask)

    # 叠加预览
    overlay = img.copy()
    overlay[green_mask > 0] = (0, 255, 0)
    overlay[yellow_mask > 0] = (0, 255, 255)
    debug = cv2.addWeighted(img, 0.55, overlay, 0.45, 0)
    cv2.imwrite(str(OUT_OVERLAY), debug)

    print(f"[INFO] 样本图：{image_path}")
    print(f"[INFO] 输出：{OUT_GREEN_MASK}, {OUT_YELLOW_MASK}, {OUT_OVERLAY}")
    print()
    print("建议填入 fishing_bot.py 的阈值：")
    print(f"HSV_GREEN_LOWER = np.array({g_lo.tolist()}, dtype=np.uint8)")
    print(f"HSV_GREEN_UPPER = np.array({g_hi.tolist()}, dtype=np.uint8)")
    print(f"HSV_YELLOW_LOWER = np.array({y_lo.tolist()}, dtype=np.uint8)")
    print(f"HSV_YELLOW_UPPER = np.array({y_hi.tolist()}, dtype=np.uint8)")


if __name__ == "__main__":
    main()
