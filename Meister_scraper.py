import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from datetime import datetime

base_url = "https://expeditionmeister.com/for-sale/expedition-trucks/"
num_pages = 14  # Total number of pages

user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Safari/605.1.15",
]

all_data = []  # List to store data from all pages

for page_num in range(1, num_pages + 1):
    url = f"{base_url}{page_num}"

    try:
        headers = {"User-Agent": random.choice(user_agents)}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, "html.parser")

        listings = soup.find_all("div", class_="simple-prod")

        for listing in listings:
            try:
                title = listing.find("a", class_="title").text.strip()
                price = listing.find("div", class_="data bottom").find("div", class_="price").text.strip()
                description = listing.find("div", class_="description isList").text.strip()
                link = listing.find("a", class_="title")["href"]

                all_data.append({
                    "Title": title,
                    "Price": price,
                    "Description": description,
                    "Link": link
                })
            except AttributeError:
                print(f"Skipping listing on page {page_num} due to missing elements.")
                continue

            time.sleep(random.uniform(1, 3))

        print(f"Scraped page {page_num}")

    except requests.exceptions.RequestException as e:
        print(f"Error on page {page_num}: {e}")
    except Exception as e:
        print(f"Error on page {page_num}: {e}")

df = pd.DataFrame(all_data)

# Excel output
today = datetime.today().strftime('%d-%m-%Y')
output_file = f"expedition_trucks_all_pages_{today}.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Trucks', index=False)
    worksheet = writer.sheets['Trucks']
    worksheet.auto_filter.ref = worksheet.dimensions

    for row in range(2, len(df) + 2):
        cell = worksheet[f"D{row}"]
        cell.hyperlink = cell.value
        cell.value = "Link"
        cell.style = "Hyperlink"

    columns = ["Title", "Price", "Description", "Link"]
    for col_num, column_title in enumerate(columns, 1):
        col_letter = get_column_letter(col_num)
        worksheet[f"{col_letter}1"].font = Font(bold=True)

    worksheet.column_dimensions["A"].width = 50
    worksheet.column_dimensions["B"].width = 15
    worksheet.column_dimensions["C"].width = 100
    worksheet.column_dimensions["D"].width = 10

print(f"Data from all pages saved to {output_file}")