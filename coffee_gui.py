#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

from pynput import keyboard


SCRIPT_NAME = "make_coffee_by_image.py"


class CoffeeGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Yihuan Coffee 控制面板")
        self.root.geometry("320x180")
        self.root.resizable(False, False)

        self.process: subprocess.Popen | None = None
        self.script_path = Path(__file__).resolve().parent / SCRIPT_NAME
        self.hotkey_listener: keyboard.Listener | None = None

        self.status_var = tk.StringVar(value="状态：未运行")
        self._build_ui()
        self._start_hotkey_listener()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=16, pady=14)
        frame.pack(fill=tk.BOTH, expand=True)

        title = tk.Label(frame, text="异环咖啡脚本控制", font=("Microsoft YaHei", 12, "bold"))
        title.pack(pady=(0, 10))

        status = tk.Label(frame, textvariable=self.status_var, anchor="w")
        status.pack(fill=tk.X, pady=(0, 10))

        btn_row = tk.Frame(frame)
        btn_row.pack(fill=tk.X, pady=(0, 10))

        start_btn = tk.Button(btn_row, text="开始 (F9)", width=14, command=self.start_script)
        start_btn.pack(side=tk.LEFT, padx=(0, 8))

        stop_btn = tk.Button(btn_row, text="结束 (F10)", width=14, command=self.stop_script)
        stop_btn.pack(side=tk.LEFT)

        tips = tk.Label(
            frame,
            text="全局快捷键：F9 开始，F10 结束",
            fg="#444444",
        )
        tips.pack(anchor="w")

    def _start_hotkey_listener(self) -> None:
        def on_press(key: keyboard.KeyCode) -> None:
            if key == keyboard.Key.f9:
                self.root.after(0, self.start_script)
            elif key == keyboard.Key.f10:
                self.root.after(0, self.stop_script)

        self.hotkey_listener = keyboard.Listener(on_press=on_press)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()

    def start_script(self) -> None:
        if self.process is not None and self.process.poll() is None:
            self.status_var.set("状态：已在运行")
            return

        if not self.script_path.exists():
            messagebox.showerror("错误", f"未找到脚本：{self.script_path}")
            self.status_var.set("状态：脚本不存在")
            return

        self.process = subprocess.Popen([sys.executable, str(self.script_path)])
        self.status_var.set(f"状态：运行中 (PID={self.process.pid})")

    def stop_script(self) -> None:
        if self.process is None or self.process.poll() is not None:
            self.status_var.set("状态：未运行")
            return

        self.process.terminate()
        try:
            self.process.wait(timeout=3)
        except Exception:
            self.process.kill()
        finally:
            self.status_var.set("状态：已停止")

    def on_close(self) -> None:
        self.stop_script()
        if self.hotkey_listener is not None:
            self.hotkey_listener.stop()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    CoffeeGuiApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
