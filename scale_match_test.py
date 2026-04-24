#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
from pathlib import Path
from typing import List, Tuple

import cv2
import numpy as np


def imread_unicode(path: Path) -> np.ndarray:
    data = np.fromfile(str(path), dtype=np.uint8)
    if data.size == 0:
        raise ValueError(f"文件为空或读取失败: {path}")
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"图片解码失败: {path}")
    return img


def run_scale_match(
    template_path: Path,
    target_path: Path,
    min_scale: float,
    max_scale: float,
    step: float,
    top_n: int,
) -> None:
    tpl = imread_unicode(template_path)
    target = imread_unicode(target_path)
    tpl_g = cv2.cvtColor(tpl, cv2.COLOR_BGR2GRAY)
    target_g = cv2.cvtColor(target, cv2.COLOR_BGR2GRAY)

    h, w = target_g.shape[:2]
    scales = np.arange(min_scale, max_scale + 1e-9, step)
    best: Tuple[float, float, int, int, Tuple[int, int]] = (-1.0, 0.0, 0, 0, (0, 0))
    top: List[Tuple[float, float, int, int, Tuple[int, int]]] = []

    for s in scales:
        tw = max(8, int(round(tpl_g.shape[1] * float(s))))
        th = max(8, int(round(tpl_g.shape[0] * float(s))))
        if tw > w or th > h:
            continue

        interp = cv2.INTER_AREA if s < 1.0 else cv2.INTER_LINEAR
        rs = cv2.resize(tpl_g, (tw, th), interpolation=interp)
        res = cv2.matchTemplate(target_g, rs, cv2.TM_CCOEFF_NORMED)
        _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(res)
        row = (float(max_val), float(s), tw, th, max_loc)

        top.append(row)
        top.sort(reverse=True, key=lambda x: x[0])
        top = top[:top_n]

        if max_val > best[0]:
            best = row

    print(f"template={template_path}")
    print(f"target={target_path}")
    print(
        f"best_score={best[0]:.6f} best_scale={best[1]:.2f} "
        f"size={best[2]}x{best[3]} loc={best[4]}"
    )
    print(f"top{top_n}:")
    for item in top:
        print(
            f"score={item[0]:.6f} scale={item[1]:.2f} "
            f"size={item[2]}x{item[3]} loc={item[4]}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="多尺度模板匹配测试脚本")
    parser.add_argument("--template", required=True, help="模板图片路径")
    parser.add_argument("--target", required=True, help="目标图片路径")
    parser.add_argument("--min-scale", type=float, default=0.2, help="最小缩放")
    parser.add_argument("--max-scale", type=float, default=2.0, help="最大缩放")
    parser.add_argument("--step", type=float, default=0.01, help="缩放步长")
    parser.add_argument("--top-n", type=int, default=10, help="输出前N名")
    args = parser.parse_args()

    run_scale_match(
        template_path=Path(args.template),
        target_path=Path(args.target),
        min_scale=args.min_scale,
        max_scale=args.max_scale,
        step=args.step,
        top_n=args.top_n,
    )


if __name__ == "__main__":
    main()
