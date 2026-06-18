; 检测数据统计工具 v2.2.14 安装程序
; 使用 Inno Setup 编译: ISCC.exe build_installer.iss

#define MyAppName "检测数据统计工具"
#define MyAppVersion "2.2.14"
#define MyAppPublisher "HotLL"
#define MyAppURL "https://github.com/HotLL233/hot"
#define MyAppExeName "检测数据统计工具_v2.2.14.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
; 安装到用户目录，避免需要管理员权限（pywebview/tkinter 在管理员模式下可能异常）
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=检测数据统计工具_v2.2.14_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
; 不需要管理员权限 — 管理员模式下 pywebview WebView2 窗口可能无法正常创建
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupIconFile=app_icon.ico

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："; Flags: checkedonce

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\{#MyAppName}\logs"
Type: files; Name: "{localappdata}\{#MyAppName}\export_config.json"
Type: dirifempty; Name: "{localappdata}\{#MyAppName}"
