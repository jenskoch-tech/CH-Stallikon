import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from datetime import datetime
import re

# Base URL
BASE_URL = "https://www.tutti.ch"
SEARCH_URL = "https://www.tutti.ch/de/q/motorraeder/Ak8CrbW90b3JjeWNsZXOUwMDAwA?sorting=newest&page="

# Initialize list to store data
data = []

# Scrape all 100 pages
for page in range(1, 101):
    print(f"Scraping page {page}...")
    try:
        response = requests.get(SEARCH_URL + str(page), timeout=10)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching page {page}: {e}")
        continue  # Skip to next page if request fails

    soup = BeautifulSoup(response.text, 'html.parser')
    listings = soup.find_all('div', {'data-private-srp-listing-item-id': True})

    if not listings:
        print("No more listings found. Stopping.")
        break  # Stop if no more listings found

    for listing in listings:
        ad_link_tag = listing.find('a', class_='mui-style-blugjv')
        ad_link = BASE_URL + ad_link_tag['href'] if ad_link_tag else "N/A"

        # Extract title
        title_tag = listing.find('div', class_='MuiBox-root mui-style-1haxbqe')
        title = title_tag.get_text(strip=True) if title_tag else "N/A"

        # Extract price using div tag -2 method
        price = "N/A"
        div_tags = listing.find_all('div')
        if len(div_tags) >= 2:
            price_div = div_tags[-2]
            price_span = price_div.find('span')
            if price_span:
                price_text = price_span.get_text(strip=True)
                price_text = price_text.replace("'", "").replace(".-", "")
                if re.match(r'^\d+$', price_text):
                    price = int(price_text)

        # If primary method fails, use backup method
        if price == "N/A":
            price_tags = listing.find_all('span')
            for tag in price_tags:
                text = tag.get_text(strip=True)
                if re.match(r"^\d{1,3}('?\d{3})?\.-$", text):  # Matches formats like 1'950.- or 12'999.-
                    price = text.replace("'", "").replace(".-", "")  # Normalize price
                    break  # Stop at first valid price

        # Extract listing description
        description = "N/A"
        description_tag = listing.find('div', class_='MuiBox-root mui-style-xe4gv6')
        if description_tag:
            description = description_tag.get_text(strip=True)

        data.append([title, price, description, ad_link])

    time.sleep(2)  # Delay to avoid getting blocked

# Create DataFrame with specified column order
columns = ["Title", "Price", "Description", "Link"]
df = pd.DataFrame(data, columns=columns)

# Save to Excel with clickable links, today's date in filename, and formatting
today = datetime.today().strftime('%d-%m-%Y')
output_file = f"motorcycle_listings_{today}.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Motorcycles', index=False)
    worksheet = writer.sheets['Motorcycles']
    worksheet.auto_filter.ref = worksheet.dimensions

    # Make links clickable and replace full URL with "Link"
    for row in range(2, len(df) + 2):
        cell = worksheet[f"D{row}"]
        cell.hyperlink = cell.value  # Make the link clickable
        cell.value = "Link"
        cell.style = "Hyperlink"

    # Format header row
    for col_num, column_title in enumerate(columns, 1):
        col_letter = get_column_letter(col_num)
        worksheet[f"{col_letter}1"].font = Font(bold=True)

    # Set column widths
    worksheet.column_dimensions["A"].width = 50    # Title
    worksheet.column_dimensions["B"].width = 9.5    # Price
    worksheet.column_dimensions["C"].width = 175    # Description
    worksheet.column_dimensions["D"].width = 10     # Link

print(f'Data extraction complete. Saved to {output_file}')
