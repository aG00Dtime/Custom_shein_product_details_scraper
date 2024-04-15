# Custom Shein Product Details Scraper
This is for my own personal use.

This PyQt5 application allows you to scrape product details from Shein URLs, including images, and save them to an Excel file.

## Features

- Drag and drop functionality for easy use.
- Multithreaded processing for faster scraping.
- Automatically calculates Guyanese Dollar (GYD) prices based on USD prices and shipping costs. (These are hard coded)
- Saves product details to an Excel file for easy access.

## Requirements

- Python 3.x
- PyQt5
- requests
- Pillow (PIL)
- BeautifulSoup4
- openpyxl
- selenium

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/aG00Dtime/Custom_shein_product_details_scraper.git
2. Install dependencies:

    ```bash
    pip install -r requirements.txt
## Usage
### Run the application:
    
    python SheinScraper.py

Drag and drop a text file containing Shein URLs into the application window.

Wait for the scraping process to complete. The progress will be displayed in the application window.

Once the scraping is done, check the output in the Excel file named YYYY-MM-DD.xlsx, where YYYY-MM-DD is the current date.

Notes
Make sure to have a stable internet connection while scraping.
The application uses Chrome WebDriver, so make sure to have Chrome installed on your system.
This application is for educational purposes only. Use it responsibly and respect the terms of service of the websites you scrape.