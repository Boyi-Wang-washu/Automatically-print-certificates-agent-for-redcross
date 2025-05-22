# 🧾 Red Cross Certificate Printing Agent

这是一个自动化脚本系统，批量登录 Red Cross 官网并下载培训证书。

## ✅ 功能

- 读取 Excel 表格学员名单
- 自动筛选符合以下条件的证书：
  - 状态为 `Valid`
  - 完成时间在 2025 年及以后
  - 不包含 `online`
- 自动勾选、选择尺寸并点击下载
- 支持中断后继续运行
- 自动生成失败学员名单（Excel）

## 🗂 项目结构
print_certificates.py # 主程序
Students to be printed ...xlsx # 学员表格（输入）
printed_emails.txt # 已打印邮箱（自动生成）
failed_certificates.xlsx # 打印失败名单（自动生成）
.gitignore # Git 忽略配置
README.md # 本说明文件

## 🚀 使用方式

```bash
pip install pandas selenium openpyxl
python print_certificates.py
运行后程序会自动批量处理。可以随时 Ctrl+C 中断，下次继续执行未完成邮箱。

由 Boyi Wang 开发，用于 Gosvea Inc. 实习自动化项目。
