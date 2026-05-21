# 检测数据统计工具

本仓库包含 **检测数据统计工具** 的多个版本以及对应的构建产物。按照约定的目录结构进行组织，便于维护和发布。

## 目录结构
```
.
├─ .git/                     # Git 仓库元数据
├─ .gitignore                # 忽略构建产物、日志等临时文件
├─ README.md                 # 本说明文档
├─ batch_count.log           # 运行时日志（已被 .gitignore 忽略）
├─ build/                    # PyInstaller 构建产物（已被 .gitignore 忽略）
├─ 可执行程序/               # 可直接运行的 .exe 文件（无控制台、带图标）
│   ├─ 检测数据统计工具_v1.3.exe
│   ├─ 检测数据统计工具_v1.4.2.exe
│   └─ 检测数据统计工具_v1.5.2.exe
├─ 安装程序/                 # Inno Setup 脚本与生成的安装包
│   ├─ installer_v1.4.2.iss
│   ├─ installer_v1.5.2.iss
│   ├─ 检测数据统计工具_Installer_v1.2.exe
│   ├─ 检测数据统计工具_Installer_v1.3.exe
│   ├─ 检测数据统计工具_Installer_v1.4.2.exe
│   └─ 检测数据统计工具_Installer_v1.5.2.exe
├─ 源码/                     # 所有源码文件（.py）
│   ├─ 检测数据统计工具_v1.2.py
│   ├─ 检测数据统计工具_v1.3.py
│   ├─ 检测数据统计工具_v1.4.2.py
│   └─ 检测数据统计工具_v1.5.2.py
└─ 检测数据统计工具_v1.5.2.spec   # PyInstaller 规格文件（已被 .gitignore 忽略）
```

## 使用说明
- **运行可执行文件**：进入 `可执行程序/`，双击对应版本的 `.exe` 即可使用（无控制台窗口、已带统一图标 `D:\图片\1.ico`）。
- **安装程序**：在 `安装程序/` 中双击对应的 `检测数据统计工具_Installer_*.exe`，会将工具安装到系统默认目录并创建快捷方式。
- **源码**：如需修改或编译新版本，请在 `源码/` 中编辑对应的 `.py` 文件，然后使用 `pyinstaller` 重新打包，或使用 Inno Setup 重新生成安装包。

## 版本发布流程（参考）
1. **修改源码** → 保存为新版本文件（如 `检测数据统计工具_v1.6.0.py`）。
2. **使用 PyInstaller** 打包：
   ```bash
   pyinstaller --onefile --windowed --icon "D:\图片\1.ico" \
               --name 检测数据统计工具_v1.6.0 检测数据统计工具_v1.6.0.py
   ```
3. **更新 Inno Setup 脚本**（`installer_v1.6.0.iss`），确保 `SetupIconFile` 与快捷方式指向新 exe。
4. **编译安装程序**：
   ```bash
   "D:\APP\Inno Setup 6\ISCC.exe" installer_v1.6.0.iss
   ```
5. **提交并推送**：`git add . && git commit -m "Add v1.6.0" && git push`。

---
*此仓库采用统一的目录结构，以便快速定位源码、可执行文件和安装程序。*