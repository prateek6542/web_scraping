# University Web Scraping Script

## Objective

This script is designed to extract and organize data from a university website, including information about courses and scholarships, and populate it into a pre-defined Excel template. All data is dynamically scraped, and missing fields are populated with "Not Available."

## Features

1. Dynamic Scraping: Scrapes university data, courses, and scholarship information from the website.
2. Excel Output: Populates the scraped data into an Excel file using the template provided.
3. Error Handling: Fields where data is unavailable are filled with "Not Available."
4. No Hardcoding: Data is extracted dynamically, ensuring flexibility and scalability.

## Prerequisites

To run this project, you will need to have the following installed:
Python 3.x
### Required Libraries:
requests
beautifulsoup4
openpyxl

You can install the required libraries by running the following command:

pip install requests beautifulsoup4 openpyxl

## File Structure
The project contains the following files:

web_scraper.py: The main Python script that handles scraping and Excel population.
Web Scraper Intern Data Fields.xlsx: The Excel template to which data is added dynamically.

## How to Run the Script
1. Clone the Repository
--First, clone the repository to your local machine using Git:

git clone https://github.com/yourusername/university-web-scraper.git

cd university-web-scraper

2. Place the Excel Template
--Ensure that your Excel template file (Web Scraper Intern Data Fields.xlsx) is located in the scraping folder of the project directory. If you already have the template file in the correct location, you don't need to modify anything.

3. Run the Script
--To run the script, open a terminal or command prompt in the project directory and execute the following command:

python web_scraper.py
This will run the web scraper and generate an updated Excel file with the required data.

4. Output
--The script will generate a new Excel file with the name Updated_Web_Scraper_Intern_Data.xlsx containing the populated data. This file will be saved in the scraping folder.

## Script Explanation

1. requests: Used to send HTTP requests to the university's website.
2. BeautifulSoup: Used to parse the HTML content and extract required data.
3. openpyxl: Used to interact with the Excel file, write the scraped data to the appropriate fields, and save the results.
4. Error Handling: If any data is missing on the website, the script populates "Not Available" in the corresponding Excel field.
