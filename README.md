# EPO Patent Status Scraper

This Python script is a web scraper that is specifically designed to retrieve patent statuses from `register.epo.org`.

## Requirements

- Python 3.7+
- Selenium WebDriver
- openpyxl

## Installation

1. Make sure Python 3 is installed. You can download it from [here](https://www.python.org/downloads/).
2. Clone the repository:
   ```bash
   git clone https://github.com/lancer1911/epo-patent-status-scraper.git
   ```
3. Navigate to the cloned directory:
   ```bash
   cd epo-patent-status-scraper
   ```
4. Install the required Python libraries using pip:
   ```bash
   pip install -r requirements.txt
   ```
5. Download the [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads) that matches your installed Chrome version and place it in the same directory as the script.

## Usage

1. Run the script:
   ```bash
   python epo-status-scraper.py
   ```
2. You will be prompted to select an Excel file that contains the EP patent numbers. Make sure the file and patent numbers are ready before running the script.
3. You will be asked to specify the column that contains the patent numbers (e.g., 'C').
4. You will be asked to specify the row number to start from (normally 2, to avoid overwriting the header).
5. The script will then scrape the status of each patent and write it into a new column in the same Excel file.

## Notes

- Be aware that scraping too many times in a short period may result in your IP being blocked by the website.
- Always double-check the scraped data for any anomalies.
- This script is intended for educational purposes only. Please use responsibly and ensure all actions comply with the website's terms of service.

## Troubleshooting

- If the script can't find the ChromeDriver, make sure the driver is in the same directory as the script and that the path is correct.
- If the script can't open the Excel file, make sure the file is not open in another program.
