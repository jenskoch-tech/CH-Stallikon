import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from datetime import datetime

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

        # Extract price (modified)
        price = "N/A"
        price_tags = listing.find_all('span', class_='MuiTypography-root MuiTypography-body1 mui-style-1yf92kr')

        if price_tags:
            price = price_tags[-1].get_text(strip=True)  # Get the last span with the class

        data.append([title, price, ad_link])

    time.sleep(2)  # Delay to avoid getting blocked

# Create DataFrame
columns = ["Title", "Price", "Ad Link"]
df = pd.DataFrame(data, columns=columns)

# Save to Excel with clickable links and today's date in filename
today = datetime.today().strftime('%d-%m-%Y')  # Get today's date in dd-mm-yyyy format
output_file = f"motorcycle_listings_{today}.xlsx"  # Add date to filename

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Motorcycles', index=False)
    worksheet = writer.sheets['Motorcycles']
    worksheet.auto_filter.ref = worksheet.dimensions  # Add filter row

    # Make links clickable
    for row in range(2, len(df) + 2):
        cell = worksheet[f"C{row}"]
        cell.hyperlink = df.iloc[row-2, 2]
        cell.style = "Hyperlink"

    # Format header row
    for col_num, column_title in enumerate(columns, 1):
        col_letter = get_column_letter(col_num)
        worksheet[f"{col_letter}1"].font = Font(bold=True)

print(f'Data extraction complete. Saved to {output_file}')