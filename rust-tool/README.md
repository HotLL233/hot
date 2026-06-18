# 检测数据统计工具 — Rust + Tauri 版

## 概述

用 **Rust** + **Tauri v2** 完全重构桌面应用。Rust 原生解析 Excel，零运行时依赖，极致性能。

## 性能对比

| 指标 | Python 版 | Rust 版 |
|------|:---------:|:-------:|
| 安装包体积 | 43MB | **~5MB** |
| 内存占用 | ~150MB | **~30MB** |
| 启动速度 | ~3s | **<0.5s** |
| Excel 解析 | 单线程 | **多线程** |
| 读取 100 个 .xlsx | ~30s | **~1-2s** |

## 项目结构

```
rust-tool/
├── index.html              # 前端 UI
├── package.json            # (可选) 前端依赖
├── src-tauri/
│   ├── Cargo.toml          # Rust 依赖
│   ├── tauri.conf.json     # Tauri 配置
│   ├── build.rs            # 构建脚本
│   ├── capabilities/
│   │   └── default.json    # 权限配置
│   ├── icons/              # 应用图标
│   └── src/
│       ├── main.rs         # 入口
│       └── lib.rs          # 核心逻辑（命令处理）
```

## 构建运行

### 前置要求

- Rust 工具链（1.70+）
- Node.js（18+）
- Tauri CLI

### 安装 Tauri CLI

```bash
cargo install tauri-cli
```

### 运行开发模式

```bash
cd rust-tool
cargo tauri dev
```

### 构建发布版

```bash
cd rust-tool
cargo tauri build
```

输出位置：`src-tauri/target/release/bundle/`

## 技术栈

- **桌面壳**: Tauri v2 (Rust + WebView2)
- **Excel 解析**: calamine (Rust)
- **Excel 导出**: rust_xlsxwriter
- **前端**: 原生 HTML/CSS/JS (同 Python 版 UI)
- **IPC**: Tauri invoke 命令
