from tkinter import *
from tkinter import ttk
import tkinter
from tkinter.filedialog import askopenfile
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import pandas as pd
from seleniumbase import Driver

root = Tk()
browser = Driver(uc=True)
#root.geometry('275x85')

def open_website():
    browser.get('https://www.businessregistry.gr/publicity/index')



def scrape_data():
    #Start scraping the data
    time.sleep(1)

    #click search
    try:
        browser.find_element(By.XPATH,'//*[@class="MuiBox-root css-1ofqig9"]/div[5]/div/div[3]/button').click()
    except:
        pass

    time.sleep(2)

    # Find all the div elements with the specified class
    div_elements = browser.find_elements(By.XPATH,'//*["MuiCardContent-root.css-1qw96cp"]')

    page = 1
    lst_of_dict = []

    while True:
        time.sleep(1)
        #button_class = '//*[@id="vertical-tabpanel-0"]/div/div/div/div[1]/nav/ul/li[9]/button'
       # Find and click the last li button using XPath
        
        print("\npage number: ",page)
        page+=1
        for i in range(1,11):
            if i==1:
                try:
                    browser.find_element(By.XPATH,'//*[@class="MuiCardContent-root css-1qw96cp"]/a').click() 
                except:
                    browser.find_element(By.XPATH,'//*[@class="MuiPaper-root MuiPaper-elevation MuiPaper-rounded MuiPaper-elevation1 MuiCard-root css-s18byi"][' + str(i) + ']/div/a/p').click() 
                scrape_data_details(lst_of_dict)
            else:
                try:
                    browser.find_element(By.XPATH,'//*[@class="MuiPaper-root MuiPaper-elevation MuiPaper-rounded MuiPaper-elevation1 MuiCard-root css-s18byi"][' + str(i) + ']/div/a/p').click() 
                    scrape_data_details(lst_of_dict)
                except:
                    print("no more data in this page")

        try:  
            time.sleep(1)
            button_xpath = '//*[@id="vertical-tabpanel-0"]/div/div/div/div[1]/nav/ul/li[last()]/button'
            button = browser.find_element(By.XPATH, button_xpath)
            button.click()
            time.sleep(1)

        except:
            #an den yparxoyn alles selides stamata
            print('\nno more pages')
            break

    print(lst_of_dict)
    # Convert the list of dictionaries to a DataFrame
    '''df = pd.DataFrame(lst_of_dict)

    # Save the DataFrame to an Excel file
    excel_file = 'scraped_filtered_data.xlsx'
    df.to_excel(excel_file, index=False)  # index=False to exclude row numbers'''

    excel_file = 'scraped_data.xlsx'
    # Read existing Excel file into a DataFrame
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        df = pd.DataFrame()  # Create an empty DataFrame if the file doesn't exist
    
    # Append the scraped dictionary as a new row
    try:
        df = df.append(lst_of_dict, ignore_index=True)
    except:pass
    
    # Save the updated DataFrame to the Excel file
    df.to_excel(excel_file, index=False)  # index=False to exclude row numbers

def scrape_data_details(lst_of_dict):

    time.sleep(2)

    #get data from table
    get_url = browser.current_url
    print("The current url is: "+str(get_url))

    # Navigate to the URL
    #browser.get(get_url)

    # Get the page source using Selenium
    page_source = browser.page_source

    # Parse the page source with BeautifulSoup
    soup = BeautifulSoup(page_source, 'html.parser')

    # Find all rows within the table body
    table_rows = soup.select('tbody.MuiTableBody-root tr.MuiTableRow-root')

    # Initialize a dictionary to store the extracted div text
    div_texts_dict = {}
    lst = []
    # Loop through each row and extract the text from columns 1 and 2
    i=0
    for row in table_rows:
        i+=1
        if i==1: continue #skip first row
        

        columns = row.find_all('td', class_='MuiTableCell-root')
        if len(columns) >= 2:
            column1_text = columns[0].get_text()
            column2_text = columns[1].get_text()
            if i<=12: #first table with basic info
                div_texts_dict[column1_text] = column2_text
                lst.append(column2_text)

            else: 
                if column1_text == "E-mail" or column1_text == "Τηλέφωνο":
                    div_texts_dict[column1_text] = column2_text
                    lst.append(column2_text)
                elif column1_text.startswith("56."):
                    div_texts_dict[column1_text] = column2_text
                    lst.append(column2_text)
                else:
                    pass
                
    # Print the dictionary of div texts
    print(div_texts_dict)
    lst_of_dict.append(div_texts_dict)

    #print(lst)
    browser.get('https://www.businessregistry.gr/publicity/index')

        

#GUI

user_input = tkinter.StringVar(root)
fromm = tkinter.StringVar(root)
too = tkinter.StringVar(root)

l1 = Label(root, text="1-Click to Open website")
l2 = Label(root, text="2-Enter filters and when you\n are ready click continue")

btn1 = Button(root, text ='Open', command = lambda: open_website())
btn2 = Button(root, text ='Continue', command = lambda: scrape_data())

l1.grid(row = 0, column = 0,  pady = 5)
btn1.grid(row = 0, column = 1,  pady = 5)

l2.grid(row = 1, column = 0,  pady = 5)
btn2.grid(row=1, column = 1,  pady = 20)

try:
    mainloop()
except:
    pass