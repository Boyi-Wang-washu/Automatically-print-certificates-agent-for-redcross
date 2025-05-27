# 🧾 Red Cross Certificate Printing Agent

这是一个用于批量打印美国红十字会（Red Cross）培训证书的自动化系统。程序可自动识别最新的 Excel 学员名单文件，批量搜索证书并完成筛选、勾选和下载操作，适合 Enrollware / Red Cross 平台配合使用。

---

## ✅ 功能亮点

- 📂 自动识别当前文件夹中**最新的 Excel 表格**
- 📧 依据 `First Name` 排序批量处理学员邮箱
- 🧠 智能筛选符合以下条件的证书：
  - 状态必须为 `Valid`
  - 完成时间必须为 **2025年及以后**
  - 课程名称不包含 `online`
- 📥 自动勾选 + 尺寸选择（11'' x 8.5''）+ 点击下载
- ⏯ 支持断点续打（中断后可继续）
- ❌ 打印失败记录将自动导出为 `failed_certificates.xlsx`

---

## 🗂 项目结构说明

redcross_agent/
├── print_certificates.py # 主程序（运行此文件）
├── 任意命名的学员表格.xlsx # 每周更新（程序将自动选取最新文件）
├── printed_emails.txt # 自动生成：已成功打印邮箱
├── failed_certificates.xlsx # 自动生成：未成功打印的邮箱名单
├── .gitignore # Git 忽略上传的缓存文件
├── README.md # 项目说明文件（当前）


---

## 🚀 使用方法

### 🧰 第一次准备

1. 安装 Python：https://www.python.org/
2. 安装依赖：
pip install pandas selenium openpyxl
3. 下载 ChromeDriver
放入本地路径，并修改代码中：
chrome_driver_path = r"C:\路径\chromedriver.exe"

📤 每次打印流程
1. 将你这周的学员 Excel 表格放入项目文件夹

2. 确保该 Excel 文件是“最后修改”的（或只留当前文件在目录中）

3. 打开终端，运行：
python print_certificates.py
程序将自动识别最新 Excel → 按 First Name 顺序处理 → 批量执行打印

4. 可随时按 Ctrl + C 中断，系统将保存当前进度并跳过已完成邮箱

♻️ 如何重置程序重新打印新表格
如果你想开始打印新的表格（而不是接着上次失败名单）：

✅ 执行以下操作：
1. 删除或重命名 printed_emails.txt 文件
删除它，程序将视为首次运行
或重命名为 printed_emails_5.12.txt 保存为历史记录

2. 替换新表格并重新运行程序即可完成“重置打印”

🧯 补打说明
运行结束后建议检查：

✅ printed_emails.txt：所有成功处理过的邮箱（系统自动跳过）

❌ failed_certificates.xlsx：未能成功打印的学员（可再次运行处理）

由 Boyi Wang 开发，用于 Gosvea Inc. 实习项目，提升 Red Cross 证书打印流程自动化效率。










