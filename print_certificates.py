import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import time

# === 配置 ===
chrome_driver_path = r"C:\Users\10846\Desktop\AI Agent实用工具\chromedriver-win64\chromedriver-win64\chromedriver.exe"
checkpoint_file = "printed_emails.txt"

# === 读取 Excel 并提取邮箱 ===
excel_files = [f for f in os.listdir() if f.endswith(".xlsx") and not f.startswith("~") and "failed" not in f]
if not excel_files:
    print("❌ 未找到任何可用的 Excel 文件，请将学员名单放在当前文件夹并确保为 .xlsx 格式。")
    exit()

# 默认选择最新修改的 Excel 文件
excel_file = max(excel_files, key=os.path.getmtime)
print(f"📥 当前使用的 Excel 文件：{excel_file}")
df = pd.read_excel(excel_file)
df_sorted = df.sort_values(by="Last Name")
email_list = df_sorted["Email"].dropna().tolist()

# === 跳过已完成邮箱 ===
if os.path.exists(checkpoint_file):
    with open(checkpoint_file, "r", encoding="utf-8") as f:
        printed_emails = set(line.strip() for line in f.readlines())
else:
    printed_emails = set()

# === 标记失败记录 ===
failed_emails = []

try:
    for test_email in email_list:
        if test_email in printed_emails:
            continue

        print(f"\n📨 当前处理邮箱：{test_email}")

        options = Options()
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--start-maximized')
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)

        try:
            driver.get("https://www.redcross.org/take-a-class/digital-certificate")
            email_input = wait.until(EC.element_to_be_clickable((By.ID, "dwfrm_certificate_email")))
            email_input.send_keys(test_email)
            find_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.certificatesubmit")))
            find_button.click()
            print("✅ 已点击查询按钮，等待结果...")

            result_area = wait.until(EC.presence_of_element_located((By.ID, "certificate")))
            time.sleep(2)
            print("✅ 成功加载证书列表")

            cert_blocks = driver.find_elements(By.CLASS_NAME, "result-certificate-dt")
            selected = 0

            for block in cert_blocks:
                try:
                    class_text = block.find_element(By.CLASS_NAME, "col-class").text.lower()
                    date_text = block.find_element(By.CLASS_NAME, "col-date").text.strip()
                    status_text = block.find_elements(By.CLASS_NAME, "col-status")[-1].text.lower()

                    if "online" in class_text:
                        print("❌ 跳过：课程为 online")
                        continue

                    try:
                        date_obj = datetime.strptime(date_text, "%b %d, %Y")
                    except:
                        print("⚠️ 日期格式不识别：", date_text)
                        continue

                    if date_obj.year < 2025:
                        print("❌ 跳过：日期早于2025")
                        continue

                    if "valid" not in status_text:
                        print("❌ 跳过：状态不是 Valid")
                        continue

                    checkbox = block.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                    driver.execute_script("arguments[0].click();", checkbox)
                    print("✅ 已勾选：", class_text, date_text, status_text)
                    selected += 1

                except Exception as e:
                    print("⚠️ 跳过一项（结构不标准）：", e)

            if selected == 0:
                print("⚠️ 没有符合条件的证书")
                failed_emails.append(test_email)
            else:
                try:
                    size_radio = driver.find_element(By.CSS_SELECTOR, "input[value='1185']")
                    driver.execute_script("arguments[0].click();", size_radio)
                    print("✅ 已选择尺寸 11x8.5")
                except:
                    print("⚠️ 未能选择尺寸")
                    failed_emails.append(test_email)

                try:
                    download_button = driver.find_element(By.ID, "download")
                    driver.execute_script("arguments[0].click();", download_button)
                    print("🎉 已点击下载")
                except:
                    print("⚠️ 未能点击下载按钮")
                    failed_emails.append(test_email)

            with open(checkpoint_file, "a", encoding="utf-8") as f:
                f.write(test_email + "\n")

            time.sleep(5)
            driver.quit()

        except Exception as e:
            print("❌ 程序报错：", e)
            failed_emails.append(test_email)
            driver.quit()

except KeyboardInterrupt:
    print("\n🛑 检测到用户中断（Ctrl+C），将保存进度...\n")

# === 输出未成功邮箱列表到 Excel ===
if failed_emails:
    print("\n❗未成功打印的邮箱如下：")
    for email in failed_emails:
        print(" -", email)
    failed_df = df_sorted[df_sorted["Email"].isin(failed_emails)]
    failed_df.to_excel("failed_certificates.xlsx", index=False)
    print("✅ 已导出未成功打印的学员信息至 failed_certificates.xlsx")
else:
    print("\n✅ 所有邮箱均已成功处理！")



