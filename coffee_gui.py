#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import ctypes
import json
import os
import psutil
import queue
import shutil
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import scrolledtext
from tkinter import ttk
import datetime as dt

from pynput import keyboard
import win32api


def get_app_base_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def get_repo_root_dir() -> Path:
    """定位仓库根目录（优先包含 run_business_flow.py 的目录）。"""
    base = get_app_base_dir()
    if (base / "run_business_flow.py").exists():
        return base
    if (base.parent / "run_business_flow.py").exists():
        return base.parent
    return base


def get_worker_python(repo_root: Path) -> str | None:
    """返回用于运行外部业务脚本的 Python 解释器路径。"""
    candidates = [
        repo_root / ".venv" / "Scripts" / "python.exe",
        repo_root / "venv" / "Scripts" / "python.exe",
    ]
    for path in candidates:
        if path.exists():
            return str(path)

    python_on_path = shutil.which("python")
    if python_on_path:
        return python_on_path
    py_launcher = shutil.which("py")
    if py_launcher:
        return py_launcher
    return None


def get_log_path() -> Path:
    return get_app_base_dir() / "数据记录" / "调试" / "coffee_gui" / "coffee_gui_debug.log"


def get_settings_path() -> Path:
    return get_app_base_dir() / "coffee_gui_settings.json"


def write_debug_log(message: str) -> None:
    try:
        ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        line = f"[{ts}] {message}\n"
        log_path = get_log_path()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        # 调试日志绝不能影响主流程
        pass


class CoffeeGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        write_debug_log("GUI __init__ begin")
        self.root = root
        self.root.title("Yihuan Coffee 控制面板")
        self.root.geometry("680x430")
        self.root.resizable(True, True)

        self.process: subprocess.Popen | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.log_thread: threading.Thread | None = None
        self.hotkey_listener: keyboard.Listener | None = None

        self.status_var = tk.StringVar(value="状态：未运行")
        self.mode_var = tk.StringVar(value="coffee")
        self.wait_after_start_var = tk.StringVar(value="50")
        self.fishing_rounds_var = tk.StringVar(value="100")
        self._x1_last_down = False  # 鼠标前侧键
        self._x2_last_down = False  # 鼠标后侧键
        self._loading_settings = False
        self.mode_display_map = {
            "咖啡模式（make_coffee_by_image）": "coffee",
            "大锤模式（大锤模式.py）": "hammer",
            "钓鱼模式（fishing_entry_flow.py）": "fishing",
        }
        self.mode_display_var = tk.StringVar(value="咖啡模式（make_coffee_by_image）")
        self._build_ui()
        self._load_settings()
        self.mode_var.trace_add("write", self._on_settings_changed)
        self.wait_after_start_var.trace_add("write", self._on_settings_changed)
        self.fishing_rounds_var.trace_add("write", self._on_settings_changed)
        self._sync_mode_display_from_mode_var()
        self._refresh_mode_specific_fields()
        self._start_hotkey_listener()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        write_debug_log("GUI __init__ done")

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

        mode_row = tk.Frame(frame)
        mode_row.pack(fill=tk.X, pady=(0, 10))
        tk.Label(mode_row, text="模式：").pack(side=tk.LEFT)
        self.mode_combo = ttk.Combobox(
            mode_row,
            textvariable=self.mode_display_var,
            values=list(self.mode_display_map.keys()),
            state="readonly",
            width=34,
        )
        self.mode_combo.pack(side=tk.LEFT)
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode_combo_changed)

        self.wait_row = tk.Frame(frame)
        self.wait_row.pack(fill=tk.X, pady=(0, 10))
        tk.Label(self.wait_row, text="开始营业后等待(秒)：").pack(side=tk.LEFT)
        tk.Entry(self.wait_row, width=8, textvariable=self.wait_after_start_var).pack(side=tk.LEFT)

        self.fishing_rounds_row = tk.Frame(frame)
        tk.Label(self.fishing_rounds_row, text="单次执行轮数：").pack(side=tk.LEFT)
        tk.Entry(self.fishing_rounds_row, width=8, textvariable=self.fishing_rounds_var).pack(side=tk.LEFT)

        tips = tk.Label(
            frame,
            text="全局快捷键：F9 开始，F10 结束（新流程循环运行）",
            fg="#444444",
        )
        tips.pack(anchor="w")

        log_label = tk.Label(frame, text="运行日志：", anchor="w")
        log_label.pack(fill=tk.X, pady=(8, 4))

        self.log_text = scrolledtext.ScrolledText(frame, height=12, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

    def _start_hotkey_listener(self) -> None:
        def on_press(key: keyboard.KeyCode) -> None:
            if key == keyboard.Key.f9:
                self.root.after(0, self.start_script)
            elif key == keyboard.Key.f10:
                self.root.after(0, self.stop_script)

        self.hotkey_listener = keyboard.Listener(on_press=on_press)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()
        self._schedule_log_poll()
        self._schedule_mouse_hotkey_poll()

    def _schedule_mouse_hotkey_poll(self) -> None:
        x1_down = bool(win32api.GetAsyncKeyState(0x05) & 0x8000)  # VK_XBUTTON1: 前侧键
        x2_down = bool(win32api.GetAsyncKeyState(0x06) & 0x8000)  # VK_XBUTTON2: 后侧键

        # 启动轮询：加入鼠标后侧键（按下沿）
        if x2_down and not self._x2_last_down:
            self.start_script()
        # 结束轮询：加入鼠标前侧键（按下沿）
        if x1_down and not self._x1_last_down:
            self.stop_script()

        self._x1_last_down = x1_down
        self._x2_last_down = x2_down
        self.root.after(40, self._schedule_mouse_hotkey_poll)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _load_settings(self) -> None:
        settings_path = get_settings_path()
        if not settings_path.exists():
            return
        try:
            data = json.loads(settings_path.read_text(encoding="utf-8"))
        except Exception as exc:
            write_debug_log(f"load settings failed: {exc}")
            return

        self._loading_settings = True
        try:
            mode = str(data.get("mode", "")).strip()
            if mode in {"coffee", "hammer", "fishing"}:
                self.mode_var.set(mode)

            wait_sec = data.get("wait_after_start_sec")
            if isinstance(wait_sec, (int, float)):
                self.wait_after_start_var.set(str(wait_sec))
            elif isinstance(wait_sec, str) and wait_sec.strip():
                self.wait_after_start_var.set(wait_sec.strip())

            fishing_rounds = data.get("fishing_rounds")
            if isinstance(fishing_rounds, (int, float)):
                self.fishing_rounds_var.set(str(int(fishing_rounds)))
            elif isinstance(fishing_rounds, str) and fishing_rounds.strip():
                self.fishing_rounds_var.set(fishing_rounds.strip())
        finally:
            self._loading_settings = False
            self._sync_mode_display_from_mode_var()
            self._refresh_mode_specific_fields()

    def _save_settings(self) -> None:
        settings_path = get_settings_path()
        payload = {
            "mode": self.mode_var.get().strip() or "coffee",
            "wait_after_start_sec": self.wait_after_start_var.get().strip() or "50",
            "fishing_rounds": self.fishing_rounds_var.get().strip() or "100",
        }
        try:
            settings_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as exc:
            write_debug_log(f"save settings failed: {exc}")

    def _on_settings_changed(self, *_args) -> None:
        if self._loading_settings:
            return
        self._save_settings()

    def _on_mode_combo_changed(self, _event=None) -> None:
        mode = self.mode_display_map.get(self.mode_display_var.get().strip(), "coffee")
        self.mode_var.set(mode)
        self._refresh_mode_specific_fields()

    def _sync_mode_display_from_mode_var(self) -> None:
        mode = self.mode_var.get().strip() or "coffee"
        reverse_map = {v: k for k, v in self.mode_display_map.items()}
        self.mode_display_var.set(reverse_map.get(mode, "咖啡模式（make_coffee_by_image）"))

    def _refresh_mode_specific_fields(self) -> None:
        mode = self.mode_var.get().strip() or "coffee"
        if mode == "fishing":
            if self.wait_row.winfo_manager():
                self.wait_row.pack_forget()
            if not self.fishing_rounds_row.winfo_manager():
                self.fishing_rounds_row.pack(fill=tk.X, pady=(0, 10))
        else:
            if self.fishing_rounds_row.winfo_manager():
                self.fishing_rounds_row.pack_forget()
            if not self.wait_row.winfo_manager():
                self.wait_row.pack(fill=tk.X, pady=(0, 10))

    def _schedule_log_poll(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass
        self.root.after(80, self._schedule_log_poll)

    def _stream_worker_output(self) -> None:
        if self.process is None or self.process.stdout is None:
            return
        try:
            for line in self.process.stdout:
                if line:
                    self.log_queue.put(line)
        except Exception as exc:
            write_debug_log(f"stdout stream thread error: {exc}")

    def start_script(self) -> None:
        write_debug_log("start_script called")
        if self.process is not None and self.process.poll() is None:
            self.status_var.set("状态：已在运行")
            write_debug_log("start_script ignored: already running")
            return

        repo_root = get_repo_root_dir()
        worker_mode = self.mode_var.get().strip() or "coffee"
        script_path = repo_root / "run_business_flow.py"
        if not script_path.exists():
            msg = f"[GUI] 未找到业务脚本：{script_path}\n"
            self._append_log(msg)
            self.status_var.set("状态：未运行")
            write_debug_log(f"start_script failed: missing {script_path}")
            return

        python_exe = get_worker_python(repo_root)
        if not python_exe:
            self.status_var.set("状态：启动失败")
            self._append_log("[GUI] 未找到可用 Python 解释器（请检查仓库 .venv）\n")
            write_debug_log("start_script failed: no python interpreter found")
            return

        wait_after_start_text = self.wait_after_start_var.get().strip()
        wait_after_start_sec = 0.0
        if worker_mode in {"coffee", "hammer"}:
            try:
                wait_after_start_sec = float(wait_after_start_text)
            except ValueError:
                self.status_var.set("状态：启动失败")
                self._append_log(f"[GUI] 等待秒数不是有效数字：{wait_after_start_text}\n")
                return
            if wait_after_start_sec < 0:
                self.status_var.set("状态：启动失败")
                self._append_log("[GUI] 等待秒数不能小于 0\n")
                return

        fishing_rounds = 100
        if worker_mode == "fishing":
            rounds_text = self.fishing_rounds_var.get().strip()
            try:
                fishing_rounds = int(rounds_text)
            except ValueError:
                self.status_var.set("状态：启动失败")
                self._append_log(f"[GUI] 钓鱼轮数不是有效整数：{rounds_text}\n")
                return
            if fishing_rounds <= 0:
                self.status_var.set("状态：启动失败")
                self._append_log("[GUI] 钓鱼轮数必须大于 0\n")
                return

        cmd = [
            python_exe,
            "-u",
            str(script_path),
            "--worker-mode",
            worker_mode,
            "--wait-after-start-sec",
            str(wait_after_start_sec),
        ]
        if worker_mode == "fishing":
            cmd.extend(["--fishing-rounds", str(fishing_rounds)])
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        creationflags = 0
        startupinfo = None
        if os.name == "nt":
            creationflags |= subprocess.CREATE_NO_WINDOW
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
        try:
            self.process = subprocess.Popen(
                cmd,
                cwd=str(repo_root),
                env=env,
                creationflags=creationflags,
                startupinfo=startupinfo,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
        except Exception as exc:
            write_debug_log(f"Popen failed: {type(exc).__name__}: {exc}")
            err_path = get_app_base_dir() / "数据记录" / "调试" / "coffee_gui" / "worker_error.log"
            err_path.parent.mkdir(parents=True, exist_ok=True)
            with err_path.open("a", encoding="utf-8") as f:
                f.write("\n=== worker start crash ===\n")
                f.write(traceback.format_exc())
            self.status_var.set("状态：启动失败")
            self._append_log(f"[GUI] 启动失败：{exc}\n")
            return

        self.log_thread = threading.Thread(target=self._stream_worker_output, daemon=True)
        self.log_thread.start()
        write_debug_log(f"worker started pid={self.process.pid} cmd={cmd} cwd={repo_root}")
        self.status_var.set(f"状态：运行中 (PID={self.process.pid})")
        self._append_log(f"[GUI] 当前模式：{worker_mode}\n")
        self._append_log(f"[GUI] 流程脚本：{script_path.name}\n")
        if worker_mode in {"coffee", "hammer"}:
            self._append_log(f"[GUI] 开始营业后等待：{wait_after_start_sec:.1f}s\n")
        if worker_mode == "fishing":
            self._append_log(f"[GUI] 单次执行轮数：{fishing_rounds}\n")
        self._append_log(f"\n[GUI] 脚本已启动，PID={self.process.pid}\n")

    def stop_script(self) -> None:
        write_debug_log("stop_script called")
        if self.process is None or self.process.poll() is not None:
            self.status_var.set("状态：未运行")
            write_debug_log("stop_script ignored: process not running")
            return

        self.process.terminate()
        try:
            self.process.wait(timeout=3)
        except Exception:
            try:
                self.process.kill()
            except Exception:
                pass
        finally:
            self.status_var.set("状态：已停止")
            self._append_log("[GUI] 脚本已停止\n")
            write_debug_log("worker stopped")
            self.process = None
            self._stop_residual_fishing_bot()

    def _stop_residual_fishing_bot(self) -> None:
        """
        前端停止时兜底清理 fishing_bot.py，避免主流程被结束后子进程残留。
        """
        repo_root = get_repo_root_dir()
        killed_pids: list[int] = []
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                cmdline = proc.info.get("cmdline") or []
                if not cmdline:
                    continue
                normalized_parts = [str(part).replace("\\", "/").lower() for part in cmdline if part]
                if not any(part.endswith("/fishing_bot.py") or part.endswith("fishing_bot.py") for part in normalized_parts):
                    continue
                # 尽量只清理当前仓库内的 fishing_bot，避免误杀。
                in_repo = any(str(repo_root).replace("\\", "/").lower() in part for part in normalized_parts)
                if not in_repo and not any(part == "fishing_bot.py" for part in normalized_parts):
                    continue
                proc.terminate()
                try:
                    proc.wait(timeout=2)
                except Exception:
                    proc.kill()
                killed_pids.append(int(proc.pid))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
            except Exception as exc:
                write_debug_log(f"stop residual fishing_bot error pid={getattr(proc, 'pid', 'unknown')}: {exc}")
        if killed_pids:
            self._append_log(f"[GUI] 已结束残留 fishing_bot 进程: {killed_pids}\n")
            write_debug_log(f"residual fishing_bot killed: {killed_pids}")

    def on_close(self) -> None:
        write_debug_log("GUI on_close")
        self.stop_script()
        if self.hotkey_listener is not None:
            self.hotkey_listener.stop()
        self.root.destroy()


def main() -> None:
    write_debug_log(f"main entry argv={sys.argv!r} frozen={getattr(sys, 'frozen', False)}")

    # 统一工作目录到仓库根目录，供外部业务脚本按相对路径读取资源
    try:
        os.chdir(get_repo_root_dir())
        write_debug_log(f"chdir -> {os.getcwd()}")
    except Exception as exc:
        write_debug_log(f"chdir failed: {exc}")

    if not ensure_admin():
        write_debug_log("ensure_admin returned False; exiting current process")
        return
    write_debug_log("ensure_admin returned True")

    root = tk.Tk()
    CoffeeGuiApp(root)
    write_debug_log("Tk mainloop starting")
    root.mainloop()
    write_debug_log("Tk mainloop exited")


def is_user_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def ensure_admin() -> bool:
    write_debug_log("ensure_admin check begin")
    if is_user_admin():
        write_debug_log("already admin")
        return True

    # 请求管理员权限后重启自身
    params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    if getattr(sys, "frozen", False):
        exe = sys.executable
    else:
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        exe = str(pythonw) if pythonw.exists() else sys.executable
        script = str(Path(__file__).resolve())
        params = f'"{script}" {params}'.strip()

    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, os.getcwd(), 1)
    write_debug_log(f"ShellExecuteW runas rc={rc}, exe={exe}, params={params}")
    # 非管理员进程只负责拉起提权后的新进程，当前进程必须退出，避免出现两个 GUI。
    return False


if __name__ == "__main__":
    main()
