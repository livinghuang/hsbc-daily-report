# -*- coding: utf-8 -*-
"""HSBC Daily Report —— Windows 桌面應用程式。

用法：
    HSBC Daily Report.exe          開啟圖形介面
    HSBC Daily Report.exe --run    不開視窗，直接產生今天的報表（給工作排程器用）
"""

from __future__ import annotations

import queue
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

import hsbc_config as cfg
import hsbc_core as core

APP_TITLE = "HSBC Daily Report"
APP_VERSION = "1.0.0"


# ======================= 共用：執行與記錄 =======================

def write_log(lines: list[str]) -> Path:
    """把這次執行的輸出寫成一份 log 檔，並清掉 30 天前的舊檔。"""
    cfg.LOG_DIR.mkdir(parents=True, exist_ok=True)

    cutoff = datetime.now().timestamp() - 30 * 86400
    for old in cfg.LOG_DIR.glob("run_*.log"):
        try:
            if old.stat().st_mtime < cutoff:
                old.unlink()
        except OSError:
            pass

    path = cfg.LOG_DIR / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def execute_report(settings: cfg.Settings, log):
    """跑一次報表，所有訊息都丟給 log callback。"""
    if not settings.excel_path:
        raise RuntimeError("還沒有指定資料來源 Excel 檔，請先到「設定」頁選擇。")
    return core.run_report(
        Path(settings.excel_path),
        settings.resolved_pdf_dir(),
        settings.cjk_font or None,
        log=log,
    )


# ========================= 無視窗模式 =========================

def run_headless() -> int:
    """給工作排程器呼叫：不開視窗，跑完就結束，結果寫進 log。

    這裡絕對不能讓例外往外丟 —— 打包成視窗程式後，未攔截的例外會跳出一個
    modal 對話框，而背景排程沒有人會去按「確定」，程式就會永遠卡在那裡。
    """
    lines = [f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {APP_TITLE} {APP_VERSION} 排程執行"]

    def log(message):
        lines.append(str(message))

    try:
        settings = cfg.Settings.load()
        result = execute_report(settings, log)
        log(f"完成：收盤日 {result.trade_date:%Y-%m-%d}")
        settings.save()
        code = 0
    except Exception as e:
        log(f"ERROR: {e}")
        log(traceback.format_exc())
        code = 1

    try:
        write_log(lines)
    except OSError:
        pass  # 連 log 都寫不出來也不能卡住，安靜結束讓排程器記下結束碼
    return code


# =========================== GUI ===========================

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    settings = cfg.Settings.load()

    root = tk.Tk()
    root.title(f"{APP_TITLE} {APP_VERSION}")
    root.geometry("820x580")
    root.minsize(760, 520)

    icon_path = Path(__file__).resolve().parent / "app.ico"
    if icon_path.exists():
        try:
            root.iconbitmap(str(icon_path))
        except tk.TclError:
            pass  # 非 Windows 平台不支援 .ico

    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("Title.TLabel", font=("Segoe UI", 15, "bold"))
    style.configure("Status.TLabel", font=("Segoe UI", 10))
    style.configure("Big.TButton", font=("Segoe UI", 10, "bold"), padding=8)

    log_queue: queue.Queue = queue.Queue()
    running = tk.BooleanVar(value=False)

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=12, pady=12)

    tab_main = ttk.Frame(notebook, padding=16)
    tab_schedule = ttk.Frame(notebook, padding=16)
    tab_settings = ttk.Frame(notebook, padding=16)
    tab_log = ttk.Frame(notebook, padding=16)
    notebook.add(tab_main, text="  主畫面  ")
    notebook.add(tab_schedule, text="  自動排程  ")
    notebook.add(tab_settings, text="  設定  ")
    notebook.add(tab_log, text="  執行紀錄  ")

    # ---------------------- 主畫面 ----------------------
    ttk.Label(tab_main, text="HSBC 每日庫存損益報表", style="Title.TLabel").pack(anchor="w")
    status_var = tk.StringVar(value="準備就緒")
    ttk.Label(tab_main, textvariable=status_var, style="Status.TLabel").pack(anchor="w", pady=(4, 12))

    button_bar = ttk.Frame(tab_main)
    button_bar.pack(fill="x", pady=(0, 12))
    run_button = ttk.Button(button_bar, text="立即產生報表", style="Big.TButton")
    run_button.pack(side="left")
    ttk.Button(
        button_bar, text="開啟 PDF 資料夾",
        command=lambda: cfg.open_folder(settings.resolved_pdf_dir()),
    ).pack(side="left", padx=8)
    open_pdf_button = ttk.Button(button_bar, text="開啟最新報表", state="disabled")
    open_pdf_button.pack(side="left")

    # 閒置時不顯示，避免進度條停在那裡讓人以為卡住了
    progress = ttk.Progressbar(tab_main, mode="indeterminate")

    columns = ("ticker", "price", "units", "market_value", "pnl")
    tree = ttk.Treeview(tab_main, columns=columns, show="headings", height=9)
    for key, title, width, anchor in (
        ("ticker", "股票", 90, "w"),
        ("price", "收盤價", 110, "e"),
        ("units", "單位數", 110, "e"),
        ("market_value", "市值", 150, "e"),
        ("pnl", "損益 P&L", 150, "e"),
    ):
        tree.heading(key, text=title)
        tree.column(key, width=width, anchor=anchor)
    tree.pack(fill="both", expand=True)
    tree.tag_configure("profit", foreground="#0a7d26")
    tree.tag_configure("loss", foreground="#c00000")

    summary_var = tk.StringVar(value="")
    ttk.Label(tab_main, textvariable=summary_var, style="Status.TLabel").pack(anchor="e", pady=(8, 0))

    # ---------------------- 自動排程 ----------------------
    ttk.Label(tab_schedule, text="每日自動執行", style="Title.TLabel").pack(anchor="w")
    schedule_status_var = tk.StringVar(value="查詢中…")
    ttk.Label(tab_schedule, textvariable=schedule_status_var, style="Status.TLabel").pack(
        anchor="w", pady=(4, 16)
    )

    time_row = ttk.Frame(tab_schedule)
    time_row.pack(anchor="w", pady=(0, 12))
    ttk.Label(time_row, text="每天執行時間：").pack(side="left")
    hour_var = tk.StringVar(value=settings.schedule_time.split(":")[0])
    minute_var = tk.StringVar(value=settings.schedule_time.split(":")[1])
    ttk.Spinbox(time_row, from_=0, to=23, width=4, format="%02.0f",
                textvariable=hour_var).pack(side="left")
    ttk.Label(time_row, text=" : ").pack(side="left")
    ttk.Spinbox(time_row, from_=0, to=59, width=4, format="%02.0f",
                textvariable=minute_var).pack(side="left")

    schedule_buttons = ttk.Frame(tab_schedule)
    schedule_buttons.pack(anchor="w", pady=(0, 16))

    note = (
        "說明：\n"
        "• 排程用 S4U 背景執行，不需要保持登入，也不會跳出視窗。\n"
        "• 電腦必須開著；已設定睡眠喚醒，但完全關機無法喚醒。\n"
        "• 若錯過執行時間，下次開機後會盡快補跑一次。"
    )
    ttk.Label(tab_schedule, text=note, style="Status.TLabel", justify="left").pack(
        anchor="w", pady=(8, 0)
    )

    def refresh_schedule_status():
        enabled, text = cfg.schedule_status()
        schedule_status_var.set(("已啟用：" if enabled else "") + text)

    def on_enable_schedule():
        time_text = f"{int(hour_var.get()):02d}:{int(minute_var.get()):02d}"
        try:
            message = cfg.enable_schedule(time_text)
            settings.schedule_time = time_text
            settings.schedule_enabled = True
            settings.save()
            refresh_schedule_status()
            messagebox.showinfo(APP_TITLE, message)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))

    def on_disable_schedule():
        try:
            message = cfg.disable_schedule()
            settings.schedule_enabled = False
            settings.save()
            refresh_schedule_status()
            messagebox.showinfo(APP_TITLE, message)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))

    ttk.Button(schedule_buttons, text="啟用／更新排程", style="Big.TButton",
               command=on_enable_schedule).pack(side="left")
    ttk.Button(schedule_buttons, text="關閉排程",
               command=on_disable_schedule).pack(side="left", padx=8)

    # ---------------------- 設定 ----------------------
    ttk.Label(tab_settings, text="資料來源與輸出", style="Title.TLabel").pack(anchor="w", pady=(0, 12))

    excel_var = tk.StringVar(value=settings.excel_path)
    pdf_var = tk.StringVar(value=settings.pdf_dir)
    font_var = tk.StringVar(value=settings.cjk_font)

    def add_path_row(parent, label, variable, browse):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=16).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="瀏覽…", command=browse).pack(side="left")

    def browse_excel():
        path = filedialog.askopenfilename(
            title="選擇 HSBC 資料來源 Excel",
            filetypes=[("Excel 活頁簿", "*.xlsx"), ("所有檔案", "*.*")],
        )
        if path:
            excel_var.set(path)

    def browse_pdf_dir():
        path = filedialog.askdirectory(title="選擇 PDF 輸出資料夾")
        if path:
            pdf_var.set(path)

    def browse_font():
        path = filedialog.askopenfilename(
            title="選擇中文字型檔",
            filetypes=[("字型檔", "*.ttf *.ttc"), ("所有檔案", "*.*")],
        )
        if path:
            font_var.set(path)

    add_path_row(tab_settings, "資料來源 Excel：", excel_var, browse_excel)
    add_path_row(tab_settings, "PDF 輸出資料夾：", pdf_var, browse_pdf_dir)
    add_path_row(tab_settings, "中文字型（選填）：", font_var, browse_font)

    hint = (
        "• 資料來源 Excel 全程唯讀，本程式不會寫入或修改這個檔案。\n"
        "• PDF 輸出資料夾留空時，預設存到 Excel 檔旁邊的 PDF 子資料夾。\n"
        "• 中文字型留空時，自動尋找細明體／微軟正黑體／標楷體。"
    )
    ttk.Label(tab_settings, text=hint, style="Status.TLabel", justify="left").pack(
        anchor="w", pady=(16, 0)
    )

    def save_settings():
        settings.excel_path = excel_var.get().strip()
        settings.pdf_dir = pdf_var.get().strip()
        settings.cjk_font = font_var.get().strip()
        settings.save()
        messagebox.showinfo(APP_TITLE, f"設定已儲存：\n{cfg.CONFIG_PATH}")

    settings_buttons = ttk.Frame(tab_settings)
    settings_buttons.pack(anchor="w", pady=(20, 0))
    ttk.Button(settings_buttons, text="儲存設定", style="Big.TButton",
               command=save_settings).pack(side="left")
    ttk.Button(settings_buttons, text="開啟設定資料夾",
               command=lambda: cfg.open_folder(cfg.app_data_dir())).pack(side="left", padx=8)

    # ---------------------- 執行紀錄 ----------------------
    ttk.Label(tab_log, text="執行紀錄", style="Title.TLabel").pack(anchor="w", pady=(0, 8))
    log_text = tk.Text(tab_log, wrap="word", height=20, font=("Consolas", 9))
    log_scroll = ttk.Scrollbar(tab_log, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=log_scroll.set, state="disabled")
    log_text.pack(side="left", fill="both", expand=True)
    log_scroll.pack(side="right", fill="y")

    def append_log(message):
        log_text.configure(state="normal")
        log_text.insert("end", f"{message}\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    # ---------------------- 執行流程 ----------------------
    collected: list[str] = []
    last_result: dict = {}

    def worker():
        try:
            result = execute_report(settings, lambda m: log_queue.put(("log", m)))
            log_queue.put(("done", result))
        except Exception as e:
            log_queue.put(("log", traceback.format_exc()))
            log_queue.put(("error", str(e)))

    def on_run():
        if running.get():
            return
        settings.excel_path = excel_var.get().strip()
        settings.pdf_dir = pdf_var.get().strip()
        settings.cjk_font = font_var.get().strip()
        settings.save()

        if not settings.excel_path:
            messagebox.showwarning(APP_TITLE, "請先到「設定」頁選擇資料來源 Excel 檔。")
            notebook.select(tab_settings)
            return

        running.set(True)
        collected.clear()
        run_button.configure(state="disabled")
        progress.pack(fill="x", pady=(0, 12), before=tree)
        progress.start(12)
        status_var.set("執行中…正在取得最新收盤價")
        for item in tree.get_children():
            tree.delete(item)
        append_log(f"—— {datetime.now():%Y-%m-%d %H:%M:%S} 開始執行 ——")
        threading.Thread(target=worker, daemon=True).start()

    run_button.configure(command=on_run)

    def finish(success: bool, message: str):
        running.set(False)
        run_button.configure(state="normal")
        progress.stop()
        progress.pack_forget()
        status_var.set(message)
        write_log(collected)
        if not success:
            messagebox.showerror(APP_TITLE, message)

    def poll_queue():
        try:
            while True:
                kind, payload = log_queue.get_nowait()
                if kind == "log":
                    collected.append(str(payload))
                    append_log(payload)
                elif kind == "done":
                    result = payload
                    last_result["pdf"] = result.pdf_path
                    for holding in result.holdings:
                        tree.insert(
                            "", "end",
                            values=(
                                holding["ticker"],
                                f"{holding['price']:,.2f}",
                                f"{holding['units']:,.0f}",
                                f"{holding['market_value']:,.2f}",
                                f"{holding['pnl']:,.2f}",
                            ),
                            tags=("profit" if holding["pnl"] >= 0 else "loss",),
                        )
                    summary_var.set(
                        f"庫存損益小計：{result.total_pnl:,.2f}　　"
                        f"已實現損益小計：{result.realized_pnl:,.2f}"
                    )
                    open_pdf_button.configure(
                        state="normal",
                        command=lambda p=result.pdf_path: cfg.open_file(p),
                    )
                    finish(True, f"完成：收盤日 {result.trade_date:%Y-%m-%d}　→　{result.pdf_path.name}")
                elif kind == "error":
                    finish(False, str(payload))
        except queue.Empty:
            pass
        root.after(120, poll_queue)

    refresh_schedule_status()
    if not settings.excel_path:
        status_var.set("尚未設定資料來源，請先到「設定」頁選擇 Excel 檔")
        notebook.select(tab_settings)

    root.after(120, poll_queue)
    root.mainloop()


def _trace_startup(stage: str):
    """把啟動過程寫進 startup.log。

    排程執行時沒有主控台可看，程式若卡在啟動階段（還沒寫出正式 log）就完全沒有線索，
    所以這裡每經過一個階段就補一行，事後才查得出卡在哪裡。
    """
    try:
        cfg.LOG_DIR.mkdir(parents=True, exist_ok=True)
        with open(cfg.LOG_DIR / "startup.log", "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {stage}\n")
    except OSError:
        pass


def main() -> int:
    _trace_startup(f"啟動 argv={sys.argv!r} frozen={getattr(sys, 'frozen', False)}")
    if "--run" in sys.argv[1:]:
        _trace_startup("模式：無視窗排程執行")
        code = run_headless()
        _trace_startup(f"無視窗執行結束，結束碼={code}")
        return code
    _trace_startup("模式：開啟圖形介面")
    run_gui()
    _trace_startup("圖形介面關閉")
    return 0


if __name__ == "__main__":
    sys.exit(main())
