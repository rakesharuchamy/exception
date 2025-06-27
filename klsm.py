from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import os
import datetime
import time
import smtplib
import pandas as pd
import re
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# === Setup Chrome Options ===
download_dir = r"E:\Tech Files\Exception Reports\KLSM"
os.makedirs(download_dir, exist_ok=True)

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}

chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# === Start WebDriver ===
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)

# === Open App and Login ===
driver.get("https://klsmsg-mack.solverminds.net/")
driver.maximize_window()
time.sleep(3)
driver.find_element(By.ID, "NFR_LoginForm-nfr_login_authname").send_keys("support")
driver.find_element(By.ID, "NFR_LoginForm-nfr_login_authid").send_keys("$vm#KLINE@2020")
driver.find_element(By.ID, "NFR_LoginForm-nfr_login_btnlogin").click()
print("Login attempted!")
time.sleep(5)

# === Navigate to Exception Report ===
search_box = wait.until(EC.presence_of_element_located((By.ID, "NFR_megamenu-nfr_topbar_autocomp1_input")))
search_box.clear()
search_box.send_keys("Exception Report")
wait.until(EC.visibility_of_element_located((By.ID, "NFR_megamenu-nfr_topbar_autocomp1_panel")))
search_box.send_keys(Keys.ENTER)
print("Navigated to Exception Report module.")
wait.until(EC.presence_of_element_located((By.ID, "EXP-searchBtn")))
time.sleep(5)

# === Dates ===
today = datetime.datetime.now().strftime("%Y-%m-%d")
yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")

# === Rename Downloaded File ===
def rename_latest_excel(date_string):
    for _ in range(20):
        files = [f for f in os.listdir(download_dir) if f.endswith(".xlsx") and "Exception Report" in f]
        if files:
            latest_file = max([os.path.join(download_dir, f) for f in files], key=os.path.getctime)
            renamed_file = os.path.join(download_dir, f"Exception Report - {date_string}.xlsx")
            try:
                os.rename(latest_file, renamed_file)
                print(f"Renamed file to: {renamed_file}")
                return renamed_file
            except FileExistsError:
                os.remove(renamed_file)
                os.rename(latest_file, renamed_file)
                print(f"Overwritten and renamed to: {renamed_file}")
                return renamed_file
        time.sleep(1)
    print(f"Failed to find file to rename for {date_string}")
    return None

# === Click and Download ===
def click_and_download(date_string):
    try:
        elements = wait.until(EC.visibility_of_all_elements_located(
            (By.XPATH, f"//ul[@class='ui-selectlistbox-list']/li[contains(text(), '{date_string}')]")))
        if elements:
            element_to_click = elements[0]
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//ul[@class='ui-selectlistbox-list']/li[contains(text(), '{date_string}')]"))).click()
            except:
                driver.execute_script("arguments[0].click();", element_to_click)
            print(f"Clicked date: {date_string}")
        else:
            print(f"Date {date_string} not found.")
            return None

        time.sleep(3)
        driver.find_element(By.CLASS_NAME, "nfr_toolpanel_li_icon").click()
        time.sleep(3)
        driver.find_element(By.CLASS_NAME, "icon-aggrid-excel").click()
        time.sleep(5)
        return rename_latest_excel(date_string)

    except Exception as e:
        print(f"Error in click_and_download: {e}")
        return None

# === Download reports ===
today_file_path = click_and_download(today)
yesterday_file_path = click_and_download(yesterday)

# === Extract Modules ===
def extract_modules_from_excel(file_path):
    try:
        df = pd.read_excel(file_path, header=None)
        for i in range(len(df)):
            row = df.iloc[i]
            for col_index, cell in enumerate(row):
                if str(cell).strip().lower() == "module name":
                    module_column = df.iloc[i+1:, col_index]
                    modules = module_column.dropna().astype(str).str.strip().str[:3]
                    return [mod.upper() for mod in modules if mod.isalpha() and len(mod) == 3]
        return []
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return []

modules_today = extract_modules_from_excel(today_file_path)
modules_yesterday = extract_modules_from_excel(yesterday_file_path)
combined_modules = sorted(set(modules_today + modules_yesterday))
print("Combined & Unique Modules:", combined_modules)

# === Module Owner Mapping ===
owner_df = pd.read_excel(os.path.join(download_dir, "ModuleDetails.xlsx"))
owner_df.columns = [col.strip().lower() for col in owner_df.columns]

owner_details = []
email_recipients = set()
for mod in combined_modules:
    row = owner_df[owner_df['module code'].str.upper() == mod]
    if not row.empty:
        owner = row.iloc[0]['owner name']
        email = row.iloc[0]['owner email']
        owner_details.append((mod, owner, email))
        email_recipients.add(email)
    else:
        owner_details.append((mod, 'N/A', 'N/A'))

# === HTML Table ===
def make_html_table_with_owner(details, date_range):
    rows = "".join(
        f"<tr><td>{i+1}</td><td>{mod}</td><td>{owner}</td></tr>"
        for i, (mod, owner, _) in enumerate(details)
    )
    return f"""
    <p><strong>{date_range}:</strong></p>
    <table border=\"1\" cellspacing=\"0\" cellpadding=\"5\">
        <tr><th>S.No</th><th>Module Code</th><th>Module Owner</th></tr>
        {rows}
    </table>
    """

# === Email Function ===
def is_valid_email(email):
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", str(email).strip()))

def send_email_with_attachment(subject, html_body, files, receivers, cc_recipients):
    sender_email = "rakesh.aruchamy@solverminds.sg"
    password = "wncbgchxgvblwqbw"

    cleaned_receivers = [e for e in receivers if is_valid_email(e)]
    cleaned_cc = [e for e in cc_recipients if is_valid_email(e)]

    msg = MIMEMultipart("alternative")
    msg['From'] = sender_email
    msg['To'] = ", ".join(cleaned_receivers)
    msg['CC'] = ", ".join(cleaned_cc)
    msg['Subject'] = subject
    msg.attach(MIMEText(html_body, 'html'))

    for file in files:
        if file:
            part = MIMEBase('application', 'octet-stream')
            with open(file, 'rb') as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file)}')
            msg.attach(part)

    with smtplib.SMTP('smtp-mail.outlook.com', 587) as server:
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, cleaned_receivers + cleaned_cc, msg.as_string())
        print("Email sent to:", cleaned_receivers + cleaned_cc)

# === Compose Email ===
subject = f"KLSM Production Exception Reports - {today} and {yesterday}"
date_range = f"{today} and {yesterday}"
html_body = f"""
<p>Dear Team,</p>
<p>Please find the Exception Reports attached for the dates {date_range} in KLSM PROD SHORE.</p>
{make_html_table_with_owner(owner_details, date_range)}
<p>Best regards,<br><strong>TechSupport</strong></p>
"""

# === Final Steps ===
to_emails = list(email_recipients)
cc_emails = ["rakesh.aruchamy@solverminds.sg"]

send_email_with_attachment(subject, html_body, [today_file_path, yesterday_file_path], to_emails, cc_emails)

# === Clean Up ===
def cleanup_files(file_paths):
    for file in file_paths:
        if file and os.path.exists(file):
            os.remove(file)
            print(f"Deleted file: {file}")

cleanup_files([today_file_path, yesterday_file_path])
driver.quit()