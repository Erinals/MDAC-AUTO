import os
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# 엑셀 파일 경로
excel_file_path = os.path.join(os.path.expanduser("~"), "Downloads", "MDAC.xlsx")
user_data = pd.read_excel(excel_file_path)

# 날짜 변환 함수
def convert_date(date):
    if isinstance(date, pd.Timestamp):
        return date.strftime("%d/%m/%Y")
    return date

user_data["passport_expiry"] = user_data["passport_expiry"].apply(convert_date)
user_data["dob"] = user_data["dob"].apply(convert_date)

# 고정된 값
fixed_values = {
    "accommodation": "HOTEL/MOTEL/REST HOUSE",
    "state": "JOHOR",
    "city": "JOHOR BAHRU",
    "nationality": "REPUBLIC OF KOREA",
    "country_code": "+82",
    "country_code_confirm": "+82"
}

# 사용자 입력을 받을 값
user_inputs = {}

# GUI 생성
root = tk.Tk()
root.title("MDAC 자동화")

# 사용자 입력 필드
fields = ["email", "email_confirm", "mobile_no", "moblie_confirm_no", "arrival_date",
          "departure_date", "mode_of_travel", "port_of_embarkation", "transport_no", "address", "postcode"]
entries = {}

for field in fields:
    frame = ttk.Frame(root)
    frame.pack(fill="x", padx=10, pady=2)
    ttk.Label(frame, text=field.replace("_", " ").title() + ": ").pack(side="left")
    entry = ttk.Entry(frame)
    entry.pack(side="right", fill="x", expand=True)
    entries[field] = entry

# 진행 상태 라벨
status_label = ttk.Label(root, text="진행 상태: 대기 중")
status_label.pack(pady=5)

# 오류 로그 창
log_frame = ttk.Frame(root)
log_frame.pack(fill="both", expand=True, padx=10, pady=5)
log_label = ttk.Label(log_frame, text="오류 로그:")
log_label.pack(anchor="w")

log_text = tk.Text(log_frame, height=10, state="disabled", bg="black", fg="red")
log_text.pack(fill="both", expand=True)

# 웹 자동화 실행 함수
def run_automation():
    global user_inputs
    user_inputs = {key: entry.get() for key, entry in entries.items()}

    # Chrome 옵션 설정
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    chromedriver_path = os.path.join(os.path.expanduser("~"), "Desktop", "chromedriver-win64", "chromedriver.exe")
    service = Service(chromedriver_path)

    driver = webdriver.Chrome(service=service, options=chrome_options)

    def wait_and_input(xpath, value):
        try:
            element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script("arguments[0].value = arguments[1];", element, value)
            return True
        except Exception as e:
            log_error(f"입력 실패: {xpath} -> {e}")
            return False

    def input_date(xpath, value):
        try:
            date_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script("arguments[0].removeAttribute('readonly')", date_element)
            driver.execute_script("arguments[0].value = arguments[1];", date_element, value)
            return True
        except Exception as e:
            log_error(f"날짜 입력 실패: {xpath} -> {e}")
            return False

    def log_error(message):
        log_text.config(state="normal")
        log_text.insert("end", message + "\n")
        log_text.config(state="disabled")
        log_text.see("end")

    def log_progress(message):
        log_text.config(state="normal")
        log_text.insert("end", message + "\n")
        log_text.config(state="disabled")
        log_text.see("end")

    def fill_form(user, index, total_users):
        try:
            driver.get("https://imigresen-online.imi.gov.my/mdac/register")
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="email"]')))

            # 진행 상태 업데이트 (현재 사용자 및 진행률 %)
            progress = (index + 1) / total_users * 100
            status_label.config(text=f"진행 중: {user['name']} ({index + 1}/{total_users}, {progress:.1f}%)")

            # 로그 업데이트
            log_progress(f"현재 {user['name']} 진행 중... ({index + 1}/{total_users}, {progress:.1f}%)")

            wait_and_input('//*[@id="name"]', user["name"])
            Select(driver.find_element(By.XPATH, '//*[@id="sex"]')).select_by_visible_text(user["sex"])
            input_date('//*[@id="dob"]', user["dob"])
            wait_and_input('//*[@id="email"]', user_inputs["email"])
            wait_and_input('//*[@id="confirmEmail"]', user_inputs["email_confirm"])

            try:
                nationality_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="nationality"]')))
                Select(nationality_element).select_by_visible_text(fixed_values["nationality"])
            except Exception as e:
                log_error(f"{user['name']} - 국적 선택 오류: {e}")

            wait_and_input('//*[@id="mobile"]', user_inputs["mobile_no"])
            wait_and_input('//*[@id="passNo"]', user["passport_no"])
            input_date('//*[@id="passExpDte"]', user["passport_expiry"])
            input_date('//*[@id="arrDt"]', user_inputs["arrival_date"])
            time.sleep(1)
            input_date('//*[@id="depDt"]', user_inputs["departure_date"])
            wait_and_input('//*[@id="vesselNm"]', user_inputs["transport_no"])

            try:
                Select(driver.find_element(By.XPATH, '//*[@id="trvlMode"]')).select_by_visible_text(
                    user_inputs["mode_of_travel"])
                Select(driver.find_element(By.XPATH, '//*[@id="embark"]')).select_by_visible_text(
                    user_inputs["port_of_embarkation"])
                Select(driver.find_element(By.XPATH, '//*[@id="accommodationStay"]')).select_by_visible_text(
                    fixed_values["accommodation"])
                Select(driver.find_element(By.XPATH, '//*[@id="accommodationState"]')).select_by_visible_text(
                    fixed_values["state"])
                time.sleep(2)
                Select(driver.find_element(By.XPATH, '//*[@id="accommodationCity"]')).select_by_visible_text(
                    fixed_values["city"])
            except NoSuchElementException as e:
                log_error(f"{user['name']} - 드롭다운 선택 오류: {e}")

            wait_and_input('//*[@id="accommodationAddress1"]', user_inputs["address"])
            wait_and_input('//*[@id="accommodationPostcode"]', user_inputs["postcode"])

            submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="submit"]')))
            submit_button.click()

            time.sleep(2)

        except Exception as e:
            log_error(f"{user['name']} - 폼 작성 중 오류 발생: {e}")

    total_users = len(user_data)
    for index, user in user_data.iterrows():
        fill_form(user, index, total_users)

    driver.quit()
    status_label.config(text="모든 작업 완료!")

# 실행 버튼 추가
run_button = ttk.Button(root, text="자동화 시작", command=run_automation)
run_button.pack(pady=10)

root.mainloop()
