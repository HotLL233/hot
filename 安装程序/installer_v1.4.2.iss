[Setup]
AppName=检测数据统计工具
AppVersion=1.4.2
DefaultDirName={autopf}\\检测数据统计工具
DefaultGroupName=检测数据统计工具
OutputBaseFilename=检测数据统计工具_Installer_v1.4.2
Compression=lzma
SolidCompression=yes

[Files]
Source: "d:\\桌面\\py\\dist\\检测数据统计工具_v1.4.2.exe"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "createShortcuts"; Description: "创建桌面和开始菜单快捷方式"; GroupDescription: "快捷方式"; Flags: unchecked

[Icons]
; 开始菜单快捷方式（可选）
Name: "{group}\\检测数据统计工具"; Filename: "{app}\\检测数据统计工具_v1.4.2.exe"; WorkingDir: "{app}"; IconFilename: "{app}\\检测数据统计工具_v1.4.2.exe"; Tasks: createShortcuts
; 桌面快捷方式（可选）
Name: "{commondesktop}\\检测数据统计工具.lnk"; Filename: "{app}\\检测数据统计工具_v1.4.2.exe"; WorkingDir: "{app}"; IconFilename: "{app}\\检测数据统计工具_v1.4.2.exe"; Tasks: createShortcuts
