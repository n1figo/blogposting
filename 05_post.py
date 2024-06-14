import pandas as pd
import numpy as np
import datetime

import pandas as pd
import numpy as np
import datetime

def read_and_filter_excel(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path, engine='openpyxl')

    # Get current date and calculate previous, current, and next month
    now = datetime.datetime.now()
    this_year = now.year
    this_month = now.month
    next_month = (now.month % 12) + 1
    previous_month = now.month - 1 if now.month > 1 else 12

    # Filter data for rooms with contract end dates this year and this month, next month, or last month
    df['만료일'] = pd.to_datetime(df['만료일'], errors='coerce')
    df_filtered = df[(df['만료일'].dt.year == this_year) & (df['만료일'].dt.month.isin([previous_month, this_month, next_month]))]

    # Convert non-business days to the next business day
    df_filtered.loc[:, '만료일'] = df_filtered['만료일'].apply(lambda x: np.busday_offset(np.datetime64(x, 'D'), 0, roll='forward'))

    # Calculate booking available date
    df_filtered.loc[:, 'Available Booking Date'] = df_filtered['만료일'].apply(lambda x: np.busday_offset(np.datetime64(x, 'D'), 4))

    # Filter the data further to include only contract end dates in the previous, current, and next month
    df_filtered = df_filtered[df_filtered['만료일'].dt.month.isin([previous_month, this_month, next_month])]

    return df_filtered

# 나머지 코드는 동일합니다...

def generate_blog_content(df, template_path, video_links_path):
    # Read template
    with open(template_path, 'r', encoding='utf-8') as file:
        template = file.read()

    # Read video links
    with open(video_links_path, 'r', encoding='utf-8') as file:
        video_links_content = file.readlines()
        video_links = {}
        for line in video_links_content:
            if line.strip() and '*' in line:
                room_number, video_link = line.split('*', 1)
                room_number = room_number.strip()
                video_link = video_link.strip()
                if video_link and video_link != '-':
                    video_links[room_number] = video_link

    # Get current year and next month
    now = datetime.datetime.now()
    next_month = (now.month % 12) + 1
    next_year = now.year if next_month != 1 else now.year + 1

    # Generate room list content
    room_list_4f = ""
    room_list_5f = ""
    for index, row in df.iterrows():
        room_number = str(row['호수'])
        room_number_int = int(room_number.replace('호', ''))  # 문자열에서 '호'를 제거하고 정수로 변환
        booking_date = row['Available Booking Date']
        video_link = video_links.get(room_number, "No Video Link Available")
        room_info = f"- {room_number}호 ({booking_date.strftime('%Y-%m-%d')}) : 예약가능 [링크]({video_link})"
        if room_number_int < 200:
            room_list_4f += f"{room_info}\n"
        else:
            room_list_5f += f"{room_info}\n"

    # Format blog content
    formatted_template = template.replace("[연도다음월]", f"{next_year}년 {next_month}월")
    formatted_template = formatted_template.replace("1. 4층 (괄호 안은 입실가능일)\n\n- [객실번호] (입실가능일) : 예약가능\n\n", f"1. 4층 (괄호 안은 입실가능일)\n\n{room_list_4f}\n")
    formatted_template = formatted_template.replace("2. 5층 (괄호 안은 입실가능일)\n\n- [객실번호] (입실가능일) : 예약가능\n\n", f"2. 5층 (괄호 안은 입실가능일)\n\n{room_list_5f}\n")

    # Blog post title
    blog_title = f"24년 {next_month}월 메가스테이 잠실 예약가능객실 안내"

    return blog_title, formatted_template

def save_blog_post(title, content, output_path):
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(f"{title}\n\n{content}")

# Example usage
excel_file_path = "./현황2.xlsx"
template_file_path = "./게시글템플릿.txt"
video_links_file_path = "./룸투어링크.txt"
output_file_path = "output_blog_post.txt"


# Step-by-step execution
filtered_data = read_and_filter_excel(excel_file_path)
blog_title, blog_content = generate_blog_content(filtered_data, template_file_path, video_links_file_path)
save_blog_post(blog_title, blog_content, output_file_path)
