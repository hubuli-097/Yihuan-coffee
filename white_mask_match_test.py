#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
独立白色区域模板匹配测试：
- 只使用模板里的白色区域参与匹配（mask）
- 输出所有调试图片到单独目录：数据记录/调试/white_mask_match_test/

用法：
1) 只测模板自身（验证流程）：
   .\\.venv\\Scripts\\python.exe white_mask_match_test.py

2) 指定场景图 + 模板图：
   .\\.venv\\Scripts\\python.exe white_mask_match_test.py "场景图.png" "素材/钓鱼/点击.png"
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Tuple

import cv2
import numpy as np


DEFAULT_TEMPLATE = Path("素材/钓鱼/点击.png")
OUTPUT_DIR = Path("数据记录/调试/white_mask_match_test")

# 白色阈值（HSV）：低饱和 + 高亮度
WHITE_S_MAX = 70
WHITE_V_MIN = 165


def imread_unicode(path: Path) -> np.ndarray | None:
    """兼容中文路径读取图片。"""
    try:
        data = np.fromfile(str(path), dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, cv2.IMREAD_COLOR)
    except Exception:
        return None


def imwrite_unicode(path: Path, image: np.ndarray) -> bool:
    """兼容中文路径保存图片。"""
    ok, buf = cv2.imencode(path.suffix or ".png", image)
    if not ok:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    buf.tofile(str(path))
    return True


def build_white_mask(bgr: np.ndarray, s_max: int = WHITE_S_MAX, v_min: int = WHITE_V_MIN) -> np.ndarray:
    """提取白色区域掩码。"""
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    lower = np.array([0, 0, v_min], dtype=np.uint8)
    upper = np.array([179, s_max, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=1)
    return mask


def masked_template_match(scene_bgr: np.ndarray, template_bgr: np.ndarray) -> Tuple[float, Tuple[int, int], np.ndarray, np.ndarray, np.ndarray]:
    """
    只用模板白色区域做模板匹配。
    返回：score, top_left, result_map, scene_white_mask, template_white_mask
    """
    scene_gray = cv2.cvtColor(scene_bgr, cv2.COLOR_BGR2GRAY)
    template_gray = cv2.cvtColor(template_bgr, cv2.COLOR_BGR2GRAY)

    scene_white_mask = build_white_mask(scene_bgr)
    template_white_mask = build_white_mask(template_bgr)

    # 仅将场景白色信息参与匹配，减少背景干扰
    scene_gray_white = cv2.bitwise_and(scene_gray, scene_gray, mask=scene_white_mask)
    template_gray_white = cv2.bitwise_and(template_gray, template_gray, mask=template_white_mask)

    if scene_gray_white.shape[0] < template_gray_white.shape[0] or scene_gray_white.shape[1] < template_gray_white.shape[1]:
        raise ValueError("场景图尺寸小于模板图，无法匹配。")

    result = cv2.matchTemplate(
        scene_gray_white,
        template_gray_white,
        cv2.TM_CCOEFF_NORMED,
        mask=template_white_mask,
    )
    _min_val, max_val, _min_loc, max_loc = cv2.minMaxLoc(result)
    return float(max_val), max_loc, result, scene_white_mask, template_white_mask


def run_white_match_test(scene_path: Path, template_path: Path, out_dir: Path = OUTPUT_DIR) -> None:
    """独立测试函数：执行匹配并输出调试图。"""
    scene = imread_unicode(scene_path)
    template = imread_unicode(template_path)
    if scene is None:
        raise FileNotFoundError(f"场景图读取失败：{scene_path}")
    if template is None:
        raise FileNotFoundError(f"模板图读取失败：{template_path}")

    score, top_left, result_map, scene_mask, template_mask = masked_template_match(scene, template)

    th, tw = template.shape[:2]
    x, y = top_left
    vis = scene.copy()
    cv2.rectangle(vis, (x, y), (x + tw, y + th), (0, 255, 255), 2)
    cv2.putText(
        vis,
        f"score={score:.4f} at ({x},{y})",
        (8, 24),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.7,
        (0, 255, 255),
        2,
        cv2.LINE_AA,
    )

    # 结果热力图
    norm = cv2.normalize(result_map, None, 0, 255, cv2.NORM_MINMAX)
    heat = cv2.applyColorMap(norm.astype(np.uint8), cv2.COLORMAP_JET)

    out_dir.mkdir(parents=True, exist_ok=True)
    imwrite_unicode(out_dir / "scene_white_mask.png", scene_mask)
    imwrite_unicode(out_dir / "template_white_mask.png", template_mask)
    imwrite_unicode(out_dir / "match_visual.png", vis)
    imwrite_unicode(out_dir / "match_heatmap.png", heat)

    print(f"[INFO] scene={scene_path}")
    print(f"[INFO] template={template_path}")
    print(f"[INFO] score={score:.6f}, top_left={top_left}")
    print(f"[INFO] 输出目录：{out_dir.resolve()}")
    print("[INFO] 输出文件：scene_white_mask.png, template_white_mask.png, match_visual.png, match_heatmap.png")


def main() -> None:
    if len(sys.argv) == 1:
        # 默认用模板自测，验证流程与阈值
        scene_path = DEFAULT_TEMPLATE
        template_path = DEFAULT_TEMPLATE
    elif len(sys.argv) == 2:
        scene_path = Path(sys.argv[1])
        template_path = DEFAULT_TEMPLATE
    else:
        scene_path = Path(sys.argv[1])
        template_path = Path(sys.argv[2])

    run_white_match_test(scene_path=scene_path, template_path=template_path)


if __name__ == "__main__":
    main()
