import os
import datetime
import time
import smtplib
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# === Setup paths ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
RESOURCE_FILE = os.path.join(BASE_DIR, "exception", "ModuleDetails.xlsx")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# === Chrome Options ===
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# === Start WebDriver ===
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 20)

# === Login to App ===
driver.get("https://klsmsg-mack.solverminds.net/")
time.sleep(3)
driver.find_element(By.ID, "NFR_LoginForm-nfr_login_authname").send_keys("support")
driver.find_element(By.ID, "NFR_LoginForm-nfr_login_authid").send_keys("$vm#KLINE@2020")
driver.find_element(By.ID, "NFR_LoginForm-nfr_login_btnlogin").click()
time.sleep(5)

# === Search and Open Exception Report ===
search_box = wait.until(EC.presence_of_element_located((By.ID, "NFR_megamenu-nfr_topbar_autocomp1_input")))
search_box.send_keys("Exception Report")
wait.until(EC.visibility_of_element_located((By.ID, "NFR_megamenu-nfr_topbar_autocomp1_panel")))
search_box.send_keys(Keys.ENTER)
wait.until(EC.presence_of_element_located((By.ID, "EXP-searchBtn")))
time.sleep(5)

# === Date Setup ===
today = datetime.datetime.now().strftime("%Y-%m-%d")
yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")

def rename_latest_excel(date_string):
    for _ in range(20):
        files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith(".xlsx") and "Exception Report" in f]
        if files:
            latest_file = max([os.path.join(DOWNLOAD_DIR, f) for f in files], key=os.path.getctime)
            renamed = os.path.join(DOWNLOAD_DIR, f"Exception Report - {date_string}.xlsx")
            try:
                os.rename(latest_file, renamed)
            except FileExistsError:
                os.remove(renamed)
                os.rename(latest_file, renamed)
            return renamed
        time.sleep(1)
    return None

def click_and_download(date_string):
    try:
        elements = wait.until(EC.visibility_of_all_elements_located(
            (By.XPATH, f"//ul[@class='ui-selectlistbox-list']/li[contains(text(), '{date_string}')]")))
        if elements:
            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, f"//ul[@class='ui-selectlistbox-list']/li[contains(text(), '{date_string}')]"))).click()
            except:
                driver.execute_script("arguments[0].click();", elements[0])
        else:
            return None
        time.sleep(3)
        driver.find_element(By.CLASS_NAME, "nfr_toolpanel_li_icon").click()
        time.sleep(3)
        driver.find_element(By.CLASS_NAME, "icon-aggrid-excel").click()
        time.sleep(5)
        return rename_latest_excel(date_string)
    except Exception as e:
        print(f"Error: {e}")
        return None

# === Download files ===
today_file = click_and_download(today)
yesterday_file = click_and_download(yesterday)

# === Module Extraction ===
def extract_modules(file_path):
    try:
        df = pd.read_excel(file_path, header=None)
        for i in range(len(df)):
            row = df.iloc[i]
            for col_index, cell in enumerate(row):
                if str(cell).strip().lower() == "module name":
                    module_column = df.iloc[i+1:, col_index]
                    modules = module_column.dropna().astype(str).str.strip().str[:3]
                    return [m.upper() for m in modules if m.isalpha() and len(m) == 3]
        return []
    except Exception as e:
        print(f"Module extraction failed: {e}")
        return []

modules = sorted(set(extract_modules(today_file) + extract_modules(yesterday_file)))

# === Module Owners Mapping ===
owner_df = pd.read_excel(RESOURCE_FILE)
owner_df.columns = [col.strip().lower() for col in owner_df.columns]

owner_details = []
recipients = set()
for mod in modules:
    row = owner_df[owner_df['module code'].str.upper() == mod]
    if not row.empty:
        owner = row.iloc[0]['owner name']
        email = row.iloc[0]['owner email']
        owner_details.append((mod, owner, email))
        recipients.add(email)
    else:
        owner_details.append((mod, "N/A", "N/A"))

# === Email Formatting ===
def is_valid_email(email):
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", str(email).strip()))

def send_email(subject, body, files, receivers, cc_list):
    sender = "rakesh.aruchamy@solverminds.sg"
    password = "wncbgchxgvblwqbw"
    to_emails = [e for e in receivers if is_valid_email(e)]
    cc_emails = [e for e in cc_list if is_valid_email(e)]

    msg = MIMEMultipart("alternative")
    msg['From'] = sender
    msg['To'] = ", ".join(to_emails)
    msg['CC'] = ", ".join(cc_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    for file in files:
        if file:
            with open(file, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file)}')
                msg.attach(part)

    with smtplib.SMTP('smtp-mail.outlook.com', 587) as server:
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, to_emails + cc_emails, msg.as_string())
        print("Email sent to:", to_emails + cc_emails)

def html_table(details):
    rows = "".join(f"<tr><td>{i+1}</td><td>{mod}</td><td>{owner}</td></tr>"
                   for i, (mod, owner, _) in enumerate(details))
    return f"""
    <p>Modules involved in reports:</p>
    <table border="1" cellpadding="5" cellspacing="0">
        <tr><th>S.No</th><th>Module Code</th><th>Module Owner</th></tr>
        {rows}
    </table>
    """

# === Email Content ===
subject = f"KLSM Exception Reports - {today} & {yesterday}"
body = f"""
<p>Dear Team,</p>
<p>Please find the Exception Reports for the dates <b>{today}</b> and <b>{yesterday}</b>.</p>
{html_table(owner_details)}
<p>Regards,<br>TechSupport</p>
"""

send_email(subject, body, [today_file, yesterday_file], list(recipients), ["rakesh.aruchamy@solverminds.sg"])

# === Cleanup ===
for f in [today_file, yesterday_file]:
    if f and os.path.exists(f):
        os.remove(f)

driver.quit()
