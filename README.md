# 🧾 Red Cross Certificate Printing Agent

这是一个用于批量打印美国红十字会（Red Cross）培训证书的自动化系统。只需将学员名单 Excel 文件放入项目文件夹，程序即可自动打开 Red Cross 官网，筛选有效证书并完成勾选与下载，支持断点续打、失败导出。

---

## ✅ 功能亮点

- 📂 **自动识别学员 Excel 文件**（无需改文件名，每周只需替换 Excel 即可）
- 📧 按照 Last Name 顺序自动读取每位学员邮箱
- 🧠 筛选逻辑包括：
  - 状态必须为 `Valid`
  - 完成时间必须为 **2025年及以后**
  - 课程名称不能包含 `online`
- 🖱️ 自动执行勾选 → 尺寸选择（11 x 8.5）→ 点击下载
- 💾 自动跳过已处理邮箱（断点续打）
- ❌ 自动导出失败记录为 `failed_certificates.xlsx`

---

## 🗂 项目结构说明
redcross_agent/
├── print_certificates.py # 主程序（运行此文件）
├── 学员表格.xlsx # 任意命名，只需放在目录中即可被识别
├── printed_emails.txt # 自动生成：已成功打印邮箱记录
├── failed_certificates.xlsx # 自动生成：未成功打印的学员名单
├── .gitignore # 忽略上传的缓存和记录文件
├── README.md # 项目说明文件（当前）


---

## 🚀 使用方法

### 🧰 第一次准备

1. 安装 Python：https://www.python.org/
2. 安装依赖：

```bash
pip install pandas selenium openpyxl

下载 ChromeDriver 并放到本地，在 print_certificates.py 中设置路径：
chrome_driver_path = r"C:\路径\chromedriver.exe"

📤 每次使用流程
将你这周的学员 Excel 表格（例如 5.19-5.25.xlsx）拖进项目文件夹

双击或运行 print_certificates.py：

程序会自动选择最新的 Excel 文件作为输入

可随时 Ctrl + C 中断，后续运行会从断点继续

所有失败打印的邮箱将导出到 failed_certificates.xlsx

🧯 补打印提示
运行结束后请检查：

✅ printed_emails.txt：包含所有成功处理过的邮箱（系统自动跳过）

❌ failed_certificates.xlsx：包含处理失败的邮箱和信息（可用于手动处理或二次导入）

由 Boyi Wang 开发，用于 Gosvea Inc. 实习自动化项目，适用于 Enrollware、Red Cross 平台证书归档打印场景。




