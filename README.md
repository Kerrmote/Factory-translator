# 工厂官方文件双向翻译软件 v2.2 (稳定版)

本版本针对大规模批量翻译场景进行了深度优化，彻底解决了 GUI 卡死、生成空白文件及网络超时处理不当等核心痛点。

## 核心改进与修复

### 1. 异步多线程架构 (QThread)
- **UI 零卡死**：所有耗时操作（文件解析、API 请求、磁盘写入）均在后台线程执行，主界面始终保持流畅响应。
- **实时同步**：通过信号槽机制实时更新进度条、状态标签及已完成列表。

### 2. 健壮的网络请求机制
- **严格超时控制**：API 请求设置了 10s 连接超时及 120s 读取超时，有效防止因网络波动导致的程序挂起。
- **智能重试**：内置指数退避重试逻辑，自动处理偶发性的网络抖动。
- **异常捕获**：全面捕获并分类处理 `APITimeoutError`、`APIConnectionError` 等 SDK 异常。

### 3. 完善的文件写入保护
- **禁止空白输出**：若翻译任务失败或被取消，程序严禁写出损坏或空白的文档，确保输出目录的整洁与准确。
- **安全停止**：支持随时点击“停止”按钮，程序将安全终止当前任务并保留已成功翻译的文件。

### 4. 详细的错误审计
- **独立日志记录**：每个文件的失败原因（包括异常堆栈）都会被详细记录在 `logs/` 目录下的日志文件中，方便后续排查。
- **弹窗提示**：任务过程中若出现错误，会通过非阻塞弹窗告知用户具体受影响的文件。

## 运行环境与依赖

### 依赖列表
```bash
pip install PyQt6 python-docx pdfplumber PyMuPDF reportlab openai keyring cryptography pandas openpyxl
```

### 运行程序
```bash
python gui_app.py
```

## 打包指南 (PyInstaller)
为了确保打包后的程序能正确访问目录和依赖，请使用以下命令：
```bash
pyinstaller --onefile --windowed \
--add-data "glossary;glossary" \
--hidden-import "pandas" \
--hidden-import "openpyxl" \
--name "FactoryTranslator_v2.2" \
gui_app.py
```

---
*由 Kerrmote Yao 开发，基于 DeepSeek 深度求索提供动力。*
