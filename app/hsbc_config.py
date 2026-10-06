# -*- coding: utf-8 -*-
"""設定檔與 Windows 工作排程器的存取。"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path

APP_NAME = "HSBC Daily Report"
TASK_NAME = "HSBC Daily Report"


def app_data_dir() -> Path:
    """設定檔與 log 放的位置（Windows 走 LocalAppData，其它平台走 ~/.config）。"""
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    path = base / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


CONFIG_PATH = app_data_dir() / "config.json"
LOG_DIR = app_data_dir() / "logs"


@dataclass
class Settings:
    excel_path: str = ""
    pdf_dir: str = ""
    cjk_font: str = ""
    schedule_enabled: bool = False
    schedule_time: str = "06:30"

    @classmethod
    def load(cls) -> "Settings":
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
                return cls(**known)
            except (json.JSONDecodeError, TypeError, OSError):
                # 設定檔壞掉時不要讓程式開不起來，直接用預設值重建
                pass
        return cls()

    def save(self):
        CONFIG_PATH.write_text(
            json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def resolved_pdf_dir(self) -> Path:
        """沒指定輸出資料夾時，預設放在 xlsx 旁邊的 PDF 子資料夾。"""
        if self.pdf_dir:
            return Path(self.pdf_dir)
        if self.excel_path:
            return Path(self.excel_path).parent / "PDF"
        return app_data_dir() / "PDF"


# ==================== Windows 工作排程器 ====================

def _run_powershell(script: str) -> subprocess.CompletedProcess:
    """跑一段 PowerShell，不要彈出黑視窗。"""
    startupinfo = None
    creationflags = 0
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = subprocess.CREATE_NO_WINDOW

    return subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True, text=True, timeout=60,
        startupinfo=startupinfo, creationflags=creationflags,
    )


def executable_command() -> tuple[str, str]:
    """排程要執行的 (執行檔, 參數)。打包成 exe 就直接叫 exe，否則叫 python 跑這個專案。"""
    exe = str(Path(sys.executable).resolve())
    if getattr(sys, "frozen", False):
        return exe, "--run"
    main_py = Path(__file__).resolve().parent / "hsbc_app.py"
    return exe, f'"{main_py}" --run'


def schedule_status() -> tuple[bool, str]:
    """查排程存不存在、幾點跑。回傳 (是否已啟用, 說明文字)。"""
    if sys.platform != "win32":
        return False, "只有 Windows 支援自動排程"

    script = (
        f'$t = Get-ScheduledTask -TaskName "{TASK_NAME}" -ErrorAction SilentlyContinue; '
        'if ($t) { $t.Triggers[0].StartBoundary } else { "NONE" }'
    )
    try:
        proc = _run_powershell(script)
    except (OSError, subprocess.SubprocessError) as e:
        return False, f"查詢排程失敗：{e}"

    out = proc.stdout.strip()
    if not out or out == "NONE":
        return False, "尚未設定自動排程"
    time_part = out.split("T")[-1][:5] if "T" in out else out
    return True, f"每天 {time_part} 自動執行"


def _run_powershell_elevated(script: str) -> None:
    """用系統管理員權限（跳 UAC）跑一段 PowerShell，失敗就丟 RuntimeError。

    S4U 排程的建立／修改／刪除都需要系統管理員權限，GUI 是一般權限，
    直接跑會得到「存取被拒」。而且 PowerShell 的 cmdlet 錯誤預設不會中止，
    以前就是這樣：註冊失敗了還印出成功，排程時間根本沒改到。
    """
    # 不用 mkdtemp：Python 3.12 的 mkdtemp 會設成只有自己能讀的 ACL，
    # 提權後寫出的檔案擁有者變成 Administrators，一般權限這邊會讀不到
    tmp = Path(tempfile.gettempdir()) / f"hsbc_task_{uuid.uuid4().hex}"
    tmp.mkdir()
    script_path = tmp / "task.ps1"
    result_path = tmp / "result.txt"
    wrapped = f"""$ErrorActionPreference = 'Stop'
try {{
{script}
}} catch {{
    Set-Content -LiteralPath '{result_path}' -Value $_.Exception.Message -Encoding UTF8
    exit 1
}}
"""
    # PowerShell 5.1 要有 BOM 才會把 .ps1 當 UTF-8 讀（描述文字有中文）
    script_path.write_text(wrapped, encoding="utf-8-sig")
    launcher = (
        "$p = Start-Process powershell.exe -Verb RunAs -Wait -PassThru -WindowStyle Hidden "
        f"-ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"{script_path}\"'; "
        "exit $p.ExitCode"
    )
    try:
        proc = _run_powershell(launcher)
        if proc.returncode == 0:
            return
        try:
            detail = result_path.read_text(encoding="utf-8-sig").strip()
        except OSError:
            detail = ""
        raise RuntimeError(detail or proc.stderr.strip() or "未取得系統管理員權限（UAC 被取消？）")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _current_user() -> str:
    # 在一般權限這邊先決定帳號：UAC 若輸入別的管理員帳密，提權後的 $env:USERNAME 會變成那個人
    return f"{os.environ.get('USERDOMAIN', '')}\\{os.environ.get('USERNAME', '')}"


def enable_schedule(time_text: str) -> str:
    """建立／更新每日排程（會跳 UAC）。

    用 S4U 登入型態：不需要存密碼，使用者不用保持登入也能背景執行。
    （這一版不需要互動式 Excel，所以不必像舊版綁定「僅登入時執行」。）
    """
    if sys.platform != "win32":
        raise RuntimeError("只有 Windows 支援自動排程")

    exe, arguments = executable_command()
    work_dir = str(Path(exe).parent)
    # 排程執行時帶 --run，走無視窗模式，不開 GUI
    script = f"""
    $action = New-ScheduledTaskAction -Execute '{exe}' -Argument '{arguments}' -WorkingDirectory '{work_dir}'
    $trigger = New-ScheduledTaskTrigger -Daily -At "{time_text}"
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Hours 1)
    $principal = New-ScheduledTaskPrincipal -UserId '{_current_user()}' `
        -LogonType S4U -RunLevel Limited
    Register-ScheduledTask -TaskName "{TASK_NAME}" -Action $action -Trigger $trigger `
        -Settings $settings -Principal $principal -Force `
        -Description "每日自動更新 HSBC 庫存損益並輸出 PDF" | Out-Null
"""
    _run_powershell_elevated(script)

    enabled, status = schedule_status()
    if not enabled or time_text not in status:
        raise RuntimeError(f"排程沒有更新成功（目前狀態：{status}）")
    return f"已設定每天 {time_text} 自動執行"


def disable_schedule() -> str:
    if sys.platform != "win32":
        raise RuntimeError("只有 Windows 支援自動排程")

    enabled, _ = schedule_status()
    if enabled:
        _run_powershell_elevated(
            f'    Unregister-ScheduledTask -TaskName "{TASK_NAME}" -Confirm:$false'
        )
    return "已關閉自動排程"


def open_folder(path: Path):
    """用系統檔案總管開啟資料夾。"""
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform == "win32":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def open_file(path: Path):
    path = Path(path)
    if sys.platform == "win32":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
