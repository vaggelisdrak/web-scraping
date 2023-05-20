# web-scraping

Created an automated desktop application for a client to scrape some info from a public government website.

1) The user fills the form (the paid API key to solve reCAPTCHA, the excel file and the search-by filter)
2) Based on the info from the given excel file, the app opens the gov website and fills the form, solves the reCAPTCHA and retrieves the data
3) The scraped data are then stored to the same excel file accordingly

Tech used:

* Pandas
* Selenium
* Tkinter
* openpyxl 
* twocaptcha API
