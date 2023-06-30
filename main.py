# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import os
from datetime import datetime
import time
from openpyxl import load_workbook

# Press the green button in the gutter to run the script.

def initChrome():
    chrome_options = Options()
    #chrome_options.add_experimental_option("detach", True)
    chrome_options.add_argument('--headless=new')
    service = ChromeService(executable_path=ChromeDriverManager().install())

    wd = webdriver.Chrome(options=chrome_options, service=service)
    return wd

def download_wait(path_to_downloads):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1)
        dl_wait = False
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds


def main():

    wd = initChrome()
    wd.implicitly_wait(10)
    wd.get("https://cronometer.com/login/")
    wd.implicitly_wait(10)

    #Not yet needed, box is automatically clicked
    wd.find_element(By.NAME, "username").send_keys("cjm3130@gmail.com")
    wd.find_element(By.NAME, "password").send_keys("G9nU7GY7bTPkP4gnL5gz")
    wd.implicitly_wait(10)
    wd.find_element(By.ID, "login-button").click()
    wd.implicitly_wait(10)
    wd.find_element(By.XPATH, "/html/body/div[1]/nav/div/div/div[2]/div/div/div/ul/li[3]/a").click()
    wd.find_element(By.XPATH, "/html/body/div[1]/nav/div/div/div[2]/div/div/div/ul/li[3]/div/ul/li[1]/a").click()
    wd.find_element(By.XPATH, "/html/body/div[1]/div/main/div[3]/div/div[2]/div[1]/div[4]/div[1]/div/div[1]/div[2]/div[3]").click()
    wd.implicitly_wait(10)
    wd.find_element(By.XPATH, "/html/body/div[4]/div/table/tbody/tr[8]/td/div").click()

    wd.implicitly_wait(10)


    download_wait(r'C:\Users\cjm31\Downloads')
    file_path = r'C:\Users\cjm31\Downloads\chart.csv'
    #read csv and delete fasting column
    data = pd.read_csv(file_path)

    #remove file

    os.remove(file_path)

    column_to_delete = 'Fasting'
    data.drop(column_to_delete, axis=1, inplace=True)
    print(data)



    # Convert the DateTime column to datetime data type
    data['DateTime'] = pd.to_datetime(data['DateTime'])

    # Extract only the date component
    data['Date'] = data['DateTime'].dt.date

    # Remove the time component from the DateTime column
    data['DateTime'] = data['DateTime'].dt.strftime('%Y-%m-%d')

    column_to_delete = 'Date'
    data.drop(column_to_delete, axis=1, inplace=True)

    print(data)

    excel_file = 'Cronometer.xlsx'
    sheet_name = data.at[0, 'DateTime']
    writer = pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace')

    data.to_excel(writer, sheet_name=sheet_name)

    writer.close()





    #writer.sheets = {ws.title: ws for ws in writer.book.worksheets}



    # Write the DataFrame to the desired cell (A12)
    #data.to_excel(writer, sheet_name=sheet_name, startrow=11, startcol=0, index=False, header=False)
    #help

    # Save the Excel file
   # writer.save()


if __name__ == '__main__':
    main()
# See PyCharm help
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
