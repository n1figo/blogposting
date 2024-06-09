import os
import time
import pandas as pd
import numpy as np
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from chromedriver_py import binary_path  # chromedriver-py 패키지 사용
import getpass

# Define paths for files
download_dir = os.path.join(os.getcwd(), "downloads")
os.makedirs(download_dir, exist_ok=True)
excel_file_path = os.path.join(download_dir, "downloaded_excel_file.xlsx")
csv_file_path = os.path.join(download_dir, "converted_csv_file.csv")
template_file_path = "게시글템플릿.txt"
video_links_file_path = "룸투어링크.txt"

# Step 1: Automate the login and download process using Selenium
def download_excel_file_with_selenium(login_url, download_url, username, password):
    # Set up Chrome options
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": download_dir}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--headless')

    # Initialize the Chrome driver
    driver = webdriver.Chrome(service=ChromeService(binary_path), options=chrome_options)

    # Navigate to the login page
    driver.get(login_url)
    
    # Find the username and password fields and input the provided credentials
    driver.find_element(By.NAME, "user_id").send_keys(username)
    driver.find_element(By.NAME, "user_pass").send_keys(password)
    
    # Submit the login form
    driver.find_element(By.NAME, "user_pass").send_keys(Keys.RETURN)
    time.sleep(5)  # Wait for the login to complete

    # Navigate to the download page
    driver.get(download_url)
    time.sleep(5)  # Wait for the page to load

    # Wait for the download link to be clickable and click it
    wait = WebDriverWait(driver, 20)
    download_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/span[3]/div/div[2]/div[3]/div/div/div[2]/div/div/div/a[8]")))
    download_link.click()
    time.sleep(10)  # Wait for the file to download

    driver.quit()

# Step 2: Convert the downloaded Excel file to CSV
def convert_xlsx_to_csv(excel_file_path, csv_file_path):
    df = pd.read_excel(excel_file_path)
    df.to_csv(csv_file_path, index=False)

# Step 3: Process the data and generate blog content
def process_accommodation_data(csv_file_path, template_file_path, video_links_file_path):
    # Read CSV file
    df = pd.read_csv(csv_file_path)
    
    # Calculate booking available date
    df['Available Booking Date'] = pd.to_datetime(df['Contract End Date']).apply(lambda x: np.busday_offset(x, 4))
    
    # Read template and video links
    with open(template_file_path, 'r', encoding='utf-8') as file:
        template = file.read()
    
    with open(video_links_file_path, 'r', encoding='utf-8') as file:
        video_links_content = file.read().split('\n')
        video_links = {}
        for line in video_links_content:
            if '*' in line:
                parts = line.split()
                room_number = parts[0].strip('*')
                video_link = parts[1]
                video_links[room_number] = video_link
    
    # Get current year and next month
    now = datetime.datetime.now()
    next_month = now.month + 1 if now.month < 12 else 1
    next_year = now.year if now.month < 12 else now.year + 1
    
    # Generate room list content
    room_list_4f = ""
    room_list_5f = ""
    for index, row in df.iterrows():
        room_number = str(row['Room Number'])
        booking_date = row['Available Booking Date']
        video_link = video_links.get(room_number, "No Video Link Available")
        room_info = f"- {room_number}호 ({booking_date.strftime('%Y-%m-%d')}) : 예약가능 [링크]({video_link})"
        if int(room_number) < 200:
            room_list_4f += f"{room_info}\n"
        else:
            room_list_5f += f"{room_info}\n"
    
    # Format blog content
    formatted_template = template.replace("[연도다음월]", f"{next_year}년 {next_month}월")
    formatted_template = formatted_template.replace("1. 4층 (괄호 안은 입실가능일)\n\n- [객실번호] (입실가능일) : 예약가능\n\n", f"1. 4층 (괄호 안은 입실가능일)\n\n{room_list_4f}\n")
    formatted_template = formatted_template.replace("2. 5층 (괄호 안은 입실가능일)\n\n- [객실번호] (입실가능일) : 예약가능\n\n", f"2. 5층 (괄호 안은 입실가능일)\n\n{room_list_5f}\n")
    
    # Blog post title
    blog_title = f"24년 {next_month}월 메가스테이 잠실역점 예약가능객실 안내"
    
    return blog_title, formatted_template

# Step 4: Main function to execute the steps
def main(login_url, download_url, template_file_path, video_links_file_path):
    # Get username and password input from the user
    username = input("Enter your username: ")
    password = getpass.getpass("Enter your password: ")
    
    # Step 1: Login and download the Excel file
    download_excel_file_with_selenium(login_url, download_url, username, password)
    
    # Step 2: Convert the Excel file to CSV
    convert_xlsx_to_csv(excel_file_path, csv_file_path)
    
    # Step 3: Process the data and generate the blog content
    title, content = process_accommodation_data(csv_file_path, template_file_path, video_links_file_path)
    
    # Output the result
    print(title)
    print(content)

# Example usage
login_url = "https://www.zipsa.net/u/login"  # Update this to the actual login URL
download_url = "https://www.zipsa.net/z/lessor/index#!/tenant/tenantManage"
main(login_url, download_url, template_file_path, video_links_file_path)
