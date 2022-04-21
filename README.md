# web_scrape
This repo contains a web scraping script written in order to automate retrieving calibration data for testing equipment used in the lab. This script was written for work, and will only work when using the company's VPN.

This script utilizes Python's requests module to obtain the HTML data from the website that stores the calibration data, uses BeautifulSoup to parse the HTML data for the relevant data, and the docx module to create a table in Microsoft Word of the data.

# How to run locally
To run this script locally, you can simply download the web_scrape.py file and run:

    python web_scrape.py


