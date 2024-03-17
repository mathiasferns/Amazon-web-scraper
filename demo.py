from bs4 import BeautifulSoup
import openpyxl

# Read the HTML file
with open('Amazon.html', 'r', encoding='utf-8') as file:
    html_content = file.read()

# Parse the HTML content
soup = BeautifulSoup(html_content, 'html.parser')

# Initialize lists to store extracted data
product_names = []
product_prices = []
product_reviews = []

# Find all divs with specified class
divs = soup.find_all('div', class_='puis-card-container s-card-container s-overflow-hidden aok-relative puis-include-content-margin puis puis-v3d1skw51z63kr2ofyhe5hr2alc s-latency-cf-section puis-card-border')

# Extract information from each div
for div in divs:
    # Extract product name
    product_name = div.find('span', class_='a-size-medium a-color-base a-text-normal')
    if product_name:
        product_names.append(product_name.text)
    else:
        product_names.append(' ')

    # Extract product price
    product_price = div.find('span', class_='a-price-whole')
    if product_price:
        product_prices.append(product_price.text)
    else:
        product_prices.append(' ')

    # Extract product reviews
    product_review = div.find('div', class_='a-row a-size-small')
    if product_review:
        product_reviews.append(product_review.text)
    else:
        product_reviews.append(' ')

# Load existing Excel workbook
workbook = openpyxl.load_workbook('AmazonProducts.xlsx')
sheet = workbook.active

# Get the last row in the sheet
last_row = sheet.max_row

# Write new data to Excel sheet
for index, (name, price, review) in enumerate(zip(product_names, product_prices, product_reviews), start=last_row + 1):
    sheet[f'A{index}'] = name
    sheet[f'B{index}'] = price
    sheet[f'C{index}'] = review

# Save the updated Excel file
workbook.save('AmazonProducts.xlsx')
