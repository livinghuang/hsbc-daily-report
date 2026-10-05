; HSBC Daily Report — Windows 安裝程式
; 用 Inno Setup 6 編譯：開啟本檔後按 F9（Build → Compile）

#define AppName "HSBC Daily Report"
#define AppNameCht "HSBC 每日庫存損益報表"
#define AppVersion "1.0.0"
#define AppPublisher "HSBC Daily Report"
#define AppExeName "HSBC Daily Report.exe"

[Setup]
AppId={{7B3F2A64-9D51-4E8C-B2A7-HSBCDAILYRPT}}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppNameCht} {#AppVersion}
AppPublisher={#AppPublisher}
VersionInfoVersion={#AppVersion}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=HSBC_Daily_Report_Setup_{#AppVersion}
SetupIconFile=..\app\app.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppNameCht}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; 裝到 Program Files 需要系統管理員；設定與報表都寫在使用者資料夾，所以執行時不需要權限
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0

[Languages]
Name: "cht"; MessagesFile: "compiler:Languages\ChineseTraditional.isl"

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "附加工作："
Name: "dailytask"; Description: "每天自動產生報表（可在程式裡調整時間）"; GroupDescription: "附加工作："

[Files]
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\app\app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README_使用說明.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\{#AppNameCht}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\app.ico"
Name: "{group}\解除安裝 {#AppNameCht}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppNameCht}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\app.ico"; Tasks: desktopicon

[Run]
; 勾了「每天自動產生報表」才建立排程，預設 06:30，之後可在程式裡改
Filename: "powershell.exe"; \
    Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""$a = New-ScheduledTaskAction -Execute '{app}\{#AppExeName}' -Argument '--run' -WorkingDirectory '{app}'; $t = New-ScheduledTaskTrigger -Daily -At '06:30'; $s = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Hours 1); $p = New-ScheduledTaskPrincipal -UserId \""$env:USERDOMAIN\$env:USERNAME\"" -LogonType S4U -RunLevel Limited; Unregister-ScheduledTask -TaskName '{#AppName}' -Confirm:$false -ErrorAction SilentlyContinue; Register-ScheduledTask -TaskName '{#AppName}' -Action $a -Trigger $t -Settings $s -Principal $p -Description '每日自動更新 HSBC 庫存損益並輸出 PDF' | Out-Null"""; \
    Flags: runhidden waituntilterminated; \
    StatusMsg: "正在建立每日自動排程..."; \
    Tasks: dailytask

Filename: "{app}\{#AppExeName}"; Description: "立即開啟 {#AppNameCht}"; \
    Flags: postinstall nowait skipifsilent

[UninstallRun]
Filename: "powershell.exe"; \
    Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""Unregister-ScheduledTask -TaskName '{#AppName}' -Confirm:$false -ErrorAction SilentlyContinue"""; \
    Flags: runhidden waituntilterminated; \
    RunOnceId: "RemoveTask"

[Code]
// 產生 PDF 時要嵌入中文字型，沒有中文字型的話報表上的中文會變空白
function ChineseFontFound(): Boolean;
begin
  Result := FileExists(ExpandConstant('{win}\Fonts\mingliu.ttc'))
    or FileExists(ExpandConstant('{win}\Fonts\msjh.ttc'))
    or FileExists(ExpandConstant('{win}\Fonts\kaiu.ttf'));
end;

procedure InitializeWizard();
begin
  if not ChineseFontFound() then
    MsgBox('偵測不到中文字型（細明體 / 微軟正黑體 / 標楷體）。' + #13#10 +
      '產生 PDF 時需要嵌入中文字型，否則報表上的中文會是空白。' + #13#10#13#10 +
      '可以到「設定 → 時間與語言 → 語言」安裝繁體中文語言套件，' + #13#10 +
      '或安裝後在程式的「設定」頁自行指定一個字型檔。', mbInformation, MB_OK);
end;
