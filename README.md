# Vidwan-Portal-Data-Scraper
Web-Scraping &amp; arranging the data of Atmiya University’s faculty from Vidwan portal

## Author 
Brijraj R. Kacha | LinkedIn: [https://www.linkedin.com/in/brijraj-kacha/]

## Overview
Faculty Data Scraper is a Python tool that automatically collects and organizes faculty information from the Vidwan portal using Selenium WebDriver and BeautifulSoup. It only requires an input file with Vidwan IDs and names, and it will extract detailed information and write it into an Excel sheet.

## Features
- Automatically collects faculty data from the Vidwan portal.
- Extracts information such as Vidwan scores, articles, books, awards, designations, organizations, citations, h-index, and more.
- Searches for additional IDs (Google Scholar, ORCID, Scopus) and gathers relevant data.
- Writes all collected information into an Excel sheet, marking any missing details as "Null."

## Usage
1. **Input File Format**
   - Prepare an input Excel file containing Vidwan IDs and names in the specified format. A sample input file format is provided in the repository.
   - Note: The sample input file and output file have had identical data of faculties removed due to privacy reasons.

2. **Running the Tool**
   - Ensure you have the appropriate ChromeDriver version for your Chrome browser. You can download it from [ChromeDriver Download](https://googlechromelabs.github.io/chrome-for-testing/).
   - Execute the Python script
   - The program will read the Vidwan IDs from the input file, scrape the data from the Vidwan portal, and write the details into an output Excel sheet.

3. **Output File**
   - The output Excel file will contain detailed faculty information. If any detail is not found, "Null" will be written for that entry.

## Files in the Repository
- `Vidwan_Data Scraping.py`: The main Python script for the application.
- `input_file_format.xlsx`: Sample input file format.
- `output_sample.xlsx`: Sample output file with name, designation, and portal ID removed for privacy reasons.
- `Vidwan Project Time Lapse.mp4`: A timelapse recording of the development process.

## Notes
- Make sure to download the suitable ChromeDriver version for your Chrome browser.
- The tool dynamically scrapes data from the Vidwan portal and other linked platforms (Google Scholar, ORCID, Scopus) if available.
