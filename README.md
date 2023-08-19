# web-scraping

Created an automated desktop application for a client to scrape some info from a public government website.

1) The user fills the form (the paid API key to solve reCAPTCHA, the excel file and the search-by filter)
2) Based on the info from the given excel file, the app opens the gov website and fills the form, solves the reCAPTCHA and retrieves the data
3) The scraped data are then stored to the same excel file accordingly


Tech used:

* Pandas
* Selenium
* BeautifulSoup (for v2)
* Tkinter
* openpyxl 
* twocaptcha API

->Converted the Python app to an executable file using Pyinstaller

![image](https://github.com/vaggelisdrak/web-scraping/assets/71725114/35ecccdd-84fc-490d-8679-e62fccabd812)

