# 天眼查企业信息批量采集工具

## 1. 简介
本工具用于批量采集天眼查上的企业信息，包括：
- 统一社会信用代码
- 注册资本
- 注册时间
- 经营状态（如：存续、注销、吊销等）

**核心优势**：
- **精准搜索**：使用天眼查站内搜索，避免搜索引擎带来的同名误判。
- **自动连接**：通过连接手动开启的 Chrome 浏览器进行采集，规避部分反爬检测。
- **智能暂停**：内置防封控机制，每查询 40 条或连续失败时会自动暂停，等待人工介入。

## 2. 环境准备

### 2.1 安装依赖
确保已安装 Python 3.7+，然后运行：
```bash
pip install playwright openpyxl
playwright install chromium
```

### 2.2 启动 Chrome 调试模式（必须步骤）
采集前，必须先关闭所有 Chrome 窗口，然后用命令行启动一个开启了远程调试端口的 Chrome。

**macOS:**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 --user-data-dir="$HOME/chrome-debug-profile"
```

**Windows:**
```bash
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-debug-profile"
```

> **注意**：启动后会弹出一个新的 Chrome 窗口，请在该窗口中打开 [天眼查](https://www.tianyancha.com) 并**登录账号**。

## 3. 使用方法

1.  **准备数据**：将待查询的企业名称填入 `files/企业信息查询表.xlsx`（从第3行开始读取）。
2.  **运行脚本**：
    ```bash
    python final_scraper.py
    ```
3.  **查看结果**：采集完成后，结果保存在 `files/企业信息查询结果_final.xlsx`。

## 4. 注意事项与人工介入

### 4.1 查询限制
天眼查对连续访问有严格限制。脚本已内置保护机制：
- **每 40 条暂停**：脚本会自动暂停，提示“已连续查询 40 条”。
- **连续失败暂停**：如果连续 3 次未查到信息，脚本也会暂停。

### 4.2 如何人工介入
当终端显示 `⚠️ 检测到需要人工介入` 时：
1.  **切回浏览器**：手动刷新当前页面，或在页面上完成滑块/点击验证。
2.  **确认正常**：确保页面能正常显示企业详情。
3.  **继续运行**：回到终端，按下 **回车键 (Enter)**，脚本将继续执行。

### 4.3 数据准确性
- 脚本会比对搜索结果与目标名称的前 4 个字，若不匹配则跳过，防止抓取到错误公司。
- “已注销”或“吊销”的企业会在备注栏中高亮标注。
- 非独立法人实体（如分公司、部门）可能无法查到独立信用代码，需人工核实。

## 5. 文件结构
```
.
├── final_scraper.py          # 主程序脚本
├── files/
│   ├── 企业信息查询表.xlsx    # 输入文件
│   ├── 企业信息查询结果_final.xlsx # 输出文件
│   └── progress_final.json   # 断点续爬记录（自动生成）
└── README.md                 # 使用说明
```
