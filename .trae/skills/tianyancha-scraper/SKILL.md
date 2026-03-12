---
name: "tianyancha-scraper"
description: "天眼查企业信息采集助手。用于在新项目中快速部署爬虫脚本，并指导用户如何配置环境、手动登录和运行采集任务。"
---

# 天眼查企业信息采集助手 (tianyancha-scraper)

本 Skill 用于帮助用户快速搭建和运行“天眼查企业信息批量采集工具”。

## 功能概述
1.  **自动部署**：一键生成 `final_scraper.py` 脚本及示例 Excel 文件。
2.  **环境检查**：检查 Python 依赖 (`playwright`, `openpyxl`) 是否安装。
3.  **操作指引**：指导用户如何开启 Chrome 调试模式并手动登录。

## 使用场景
- 当用户需要批量查询企业工商信息（信用代码、注册资本、状态等）时。
- 当用户询问“如何爬取天眼查数据”时。
- 当用户想要在新项目中复用之前的采集方案时。

## 执行流程

### 1. 部署脚本文件
首先，检查当前目录下是否存在 `final_scraper.py`。如果不存在，则创建该文件，内容如下：

```python
"""
天眼查企业信息批量采集工具 (Skill 生成版)
依赖：pip install playwright openpyxl
"""
import asyncio
import json
import re
import random
from pathlib import Path
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 配置
# ============================================================
INPUT_FILE = "files/企业信息查询表.xlsx"
OUTPUT_FILE = "files/企业信息查询结果_final.xlsx"
PROGRESS_FILE = "files/progress_final.json"
CDP_PORT = 9222
TEST_LIMIT = None  # 设置为整数可限制测试数量，None 为全量

NON_LEGAL_ENTITIES = ["杭州弧途科技-用工业务部"]  # 示例非法人实体

# ============================================================
# 人工介入逻辑
# ============================================================
def wait_for_human_intervention(reason: str):
    """暂停脚本，等待人工介入"""
    print("\n" + "!" * 60)
    print(f"  ⚠️  检测到需要人工介入：{reason}")
    print("  👉 请在 Chrome 浏览器中手动完成验证（如滑块、验证码）或刷新页面")
    print("  ✅ 完成后，请回到此处按回车键继续...")
    print("!" * 60 + "\n")
    print("\a") 
    input()
    print("  ▶️  继续运行...")

# ============================================================
# 读取 Excel 中的企业列表
# ============================================================
def read_companies_from_excel(file_path):
    companies = []
    try:
        if not Path(file_path).exists():
            print(f"❌ 文件不存在: {file_path}")
            return []
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[1]:
                company_name = str(row[1]).strip()
                if company_name not in companies:
                    companies.append(company_name)
    except Exception as e:
        print(f"读取 Excel 文件失败: {e}")
        return []
    return companies

# ============================================================
# 爬取核心逻辑
# ============================================================
async def find_tianyancha_url(page, name: str) -> str:
    """通过天眼查站内搜索获取详情页链接"""
    search_url = f"https://www.tianyancha.com/search?key={name}"
    try:
        await page.goto(search_url, wait_until="domcontentloaded", timeout=20000)
        await asyncio.sleep(random.uniform(2, 3))

        if "验证" in await page.title() or await page.query_selector(".baxia-dialog") or "反爬" in await page.inner_text("body"):
            wait_for_human_intervention("触发天眼查验证/反爬拦截")
            await page.reload(wait_until="domcontentloaded")
            await asyncio.sleep(2)

        first_result = await page.query_selector(".index_list-wrap___axcs .index_name__qEdWi a")
        if not first_result:
            first_result = await page.query_selector("a[href*='company/']")
            
        if first_result:
            href = await first_result.get_attribute("href")
            text = await first_result.inner_text()
            clean_target = name.replace("（", "(").replace("）", ")").strip()
            clean_found = text.replace("（", "(").replace("）", ")").strip()
            
            if clean_target[:4] in clean_found:
                 if href and not href.startswith("http"):
                    href = "https://www.tianyancha.com" + href
                 return href
            else:
                print(f"    ⚠️ 搜索结果不匹配: 目标[{name}] vs 结果[{text}]")
                return ""
    except Exception as e:
        print(f"    站内搜索失败: {e}")
    return ""

async def extract_from_tianyancha(page, url: str, name: str) -> dict:
    result = {"name": name, "credit_code": "", "registered_capital": "",
              "registration_date": "", "status": "", "remark": ""}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=25000)
        await asyncio.sleep(random.uniform(2, 3))

        login_modal = await page.query_selector(".dialog-content, .login-dialog, [class*='login-modal']")
        if login_modal:
            close_btn = await page.query_selector(".close-btn, .dialog-close, [class*='close']")
            if close_btn:
                await close_btn.click()
                await asyncio.sleep(1)

        text = await page.inner_text("body")

        m = re.search(r"统一社会信用代码[：:\s]*([A-Z0-9]{18})", text)
        if m: result["credit_code"] = m.group(1)

        capital_patterns = [
            r"注册资本[：:\s]*([0-9,.]+\s*(?:万|亿|万美元|万港元|万人民币|万元人民币|万元|元)[人民币]*)",
            r"注册资本[：:\s]*([0-9,.]+\s*[^\s<]+)",
        ]
        for pat in capital_patterns:
            m = re.search(pat, text)
            if m:
                cap = m.group(1).strip()
                if len(cap) < 30:
                    result["registered_capital"] = cap
                    break

        for pat in [r"成立日期[：:\s]*(\d{4}-\d{2}-\d{2})", r"注册时间[：:\s]*(\d{4}-\d{2}-\d{2})", r"成立时间[：:\s]*(\d{4}-\d{2}-\d{2})"]:
            m = re.search(pat, text)
            if m:
                result["registration_date"] = m.group(1)
                break

        m = re.search(r"(存续|注销|吊销|迁出|撤销|开业|在营|登记)", text)
        if m:
            result["status"] = m.group(1)
            if result["status"] in ["注销", "吊销", "撤销"]:
                result["remark"] = f"⚠️ 企业已{result['status']}" + (" | " + result["remark"] if result["remark"] else "")

    except Exception as e:
        result["remark"] = f"提取出错: {str(e)[:60]}"
        print(f"    ❌ 提取失败: {e}")
    return result

async def query_company(page, name: str) -> dict:
    result = {"name": name, "credit_code": "", "registered_capital": "",
              "registration_date": "", "status": "", "remark": ""}
    print(f"  🔍 搜索天眼查链接...")
    url = await find_tianyancha_url(page, name)
    if not url:
        result["remark"] = "未找到天眼查页面"
        print(f"  ❌ 未找到链接")
        return result

    print(f"  🌐 访问: {url}")
    result = await extract_from_tianyancha(page, url, name)
    result["name"] = name
    if result["credit_code"]:
        print(f"  ✅ {result['credit_code']} | {result['registered_capital']} | {result['registration_date']}")
    else:
        if not result["remark"]: result["remark"] = "未提取到信用代码，请手动核查"
        print(f"  ⚠️  数据不完整")
    return result

# ============================================================
# Excel 输出
# ============================================================
def save_excel(results: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "企业信息"

    hf = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    af = PatternFill("solid", start_color="D6E4F0", end_color="D6E4F0")
    wf = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    gf = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
    bd = Border(left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
                top=Side(style="thin", color="BFBFBF"),  bottom=Side(style="thin", color="BFBFBF"))

    headers = ["序号", "企业名称", "统一社会信用代码", "注册资本", "注册时间", "经营状态", "备注"]
    widths  = [6, 40, 22, 16, 14, 10, 30]

    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        cell.fill = hf
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = bd
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 22

    row = 2
    for nm in NON_LEGAL_ENTITIES:
        for c, v in enumerate([row-1, nm, "—", "—", "—", "—", "非独立法人（部门）"], 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.font = Font(name="Arial", size=10, color="808080", italic=True)
            cell.fill = gf
            cell.alignment = Alignment(horizontal="center" if c != 2 else "left", vertical="center")
            cell.border = bd
        row += 1

    for i, r in enumerate(results):
        fill = wf if r.get("remark") else (af if i % 2 == 0 else PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF"))
        for c, v in enumerate([row-1, r["name"], r["credit_code"], r["registered_capital"],
                               r["registration_date"], r["status"], r.get("remark", "")], 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.font = Font(name="Arial", size=10)
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center" if c in [1,3,4,5,6] else "left", vertical="center")
            cell.border = bd
        ws.row_dimensions[row].height = 18
        row += 1

    Path(OUTPUT_FILE).parent.mkdir(exist_ok=True, parents=True)
    wb.save(OUTPUT_FILE)
    print(f"\n📊 Excel 已保存：{OUTPUT_FILE}，共 {row-2} 行数据")

# ============================================================
# 主流程
# ============================================================
async def main():
    print("=" * 60)
    print("  企业信息批量爬虫 (天眼查版)")
    print("=" * 60)
    print(f"正在读取 {INPUT_FILE} ...")
    
    all_companies = read_companies_from_excel(INPUT_FILE)
    if not all_companies:
        print("❌ 未能读取到企业列表，请检查 Excel 文件。")
        return

    companies_to_process = all_companies
    if TEST_LIMIT:
        companies_to_process = all_companies[:TEST_LIMIT]
        print(f"共读取 {len(all_companies)} 家，本次测试 {len(companies_to_process)} 家")
    else:
        print(f"共读取 {len(all_companies)} 家，准备全量查询")

    progress = {}
    if Path(PROGRESS_FILE).exists():
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            progress = json.load(f)
    done = set(progress.get("done", {}).keys())
    todo = [c for c in companies_to_process if c not in done]
    all_results = list(progress.get("done", {}).values())
    print(f"待查：{len(todo)} 家 / 已完成：{len(done)} 家\n")

    if not todo:
        save_excel(all_results)
        return

    print("⚠️  请确认已启动 Chrome 调试模式并登录天眼查！")
    # input("✅ 确认请按回车...") # 自动化时可注释

    try:
        async with async_playwright() as p:
            browser = await p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
            print(f"✅ 已连接 Chrome（端口 {CDP_PORT}）\n")
            contexts = browser.contexts
            page = contexts[0].pages[0] if (contexts and contexts[0].pages) else await browser.new_page()

            consecutive_failures = 0
            for i, name in enumerate(todo, 1):
                if i > 0 and i % 40 == 0:
                    wait_for_human_intervention("已连续查询 40 条，天眼查可能会限制访问")

                print(f"\n[{i}/{len(todo)}] {name}")
                result = await query_company(page, name)
                
                if not result["credit_code"] and "未找到" in result.get("remark", ""):
                    consecutive_failures += 1
                    if consecutive_failures >= 3:
                        wait_for_human_intervention("连续 3 次未找到有效信息，可能已被封控")
                        consecutive_failures = 0
                else:
                    consecutive_failures = 0

                all_results.append(result)
                progress.setdefault("done", {})[name] = result
                with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                    json.dump(progress, f, ensure_ascii=False, indent=2)

                delay = random.uniform(3, 6)
                print(f"  ⏳ 等待 {delay:.1f}s")
                await asyncio.sleep(delay)

            await browser.close()
    except Exception as e:
        print(f"\n❌ 连接 Chrome 失败: {e}")
        print("请检查：1. Chrome 是否已关闭所有窗口 2. 启动命令是否包含 --remote-debugging-port=9222")

    name_map = {r["name"]: r for r in all_results}
    ordered  = [name_map[c] for c in companies_to_process if c in name_map]
    save_excel(ordered)
    print("\n🎉 全部完成！")

if __name__ == "__main__":
    asyncio.run(main())
```

同时，确保 `files/` 目录存在，并提示用户创建 `files/企业信息查询表.xlsx`（如果不存在）。

### 2. 环境检查与安装
自动运行 `pip install playwright openpyxl` 确保依赖存在。
提示用户运行 `playwright install chromium`（如果尚未安装）。

### 3. 用户指引
输出详细的启动指引，告诉用户如何启动 Chrome：

**macOS:**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 --user-data-dir="$HOME/chrome-debug-profile"
```

**Windows:**
```bash
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-debug-profile"
```

### 4. 运行脚本
一切就绪后，提示用户运行 `python final_scraper.py`。
