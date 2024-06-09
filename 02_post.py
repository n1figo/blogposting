import os
import time
import pandas as pd
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup

# Define paths for files
excel_file_path = "downloaded_excel_file.xlsx"
csv_file_path = "converted_csv_file.csv"
template_file_path = "게시글템플릿.txt"
video_links_file_path = "룸투어링크.txt"

# Step 1: Download the Excel file using requests and BeautifulSoup
def download_excel_file(download_url):
    session = requests.Session()
    
    # Get the initial page to retrieve any necessary cookies or tokens
    response = session.get(download_url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Find the download link - adjust the selector as needed based on the actual HTML structure
    download_link = soup.select_one('a[href*=".xlsx"]')['href']
    
    # Download the Excel file
    response = session.get(download_link)
    
    # Save the file
    with open(excel_file_path, 'wb') as file:
        file.write(response.content)

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
def main(download_url, template_file_path, video_links_file_path):
    # Step 1: Download the Excel file
    download_excel_file(download_url)
    
    # Step 2: Convert the Excel file to CSV
    convert_xlsx_to_csv(excel_file_path, csv_file_path)
    
    # Step 3: Process the data and generate the blog content
    title, content = process_accommodation_data(csv_file_path, template_file_path, video_links_file_path)
    
    # Output the result
    print(title)
    print(content)

# Example usage
download_url = "https://www.zipsa.net/z/lessor/index#!/tenant/tenantManage"
main(download_url, template_file_path, video_links_file_path)
