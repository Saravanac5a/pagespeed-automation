from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

workbook = Workbook()
workbook.save("C:/Users/C207641/Downloads/pagespeed_scores.xlsx")

book = openpyxl.load_workbook("C:/Users/C207641/Downloads/pagespeed_scores.xlsx")
sheet = book.active

urls = ['https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2F','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fopen-demat-account','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fshare-market-today','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fipo','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fipo%2Foyo-ipo','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fblog%2Fgold-price-today-22-carat-and-24-carat-gold-rate','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fnews%2F3-it-stocks-to-watch-out-for-on-september-20','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fstocks%2Ftatasteel-share-price','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fderivatives','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Ftechnology%2Ftrade-station-exe','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fshare-market-today%2Fnifty-50-live','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fmutual-funds%2Famc%2Fiifl-mutual-fund%2F','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fmutual-funds%2Fcanara-robeco-bluechip-equity-fund-direct-growth%2F','https://pagespeed.web.dev/report?url=https%3A%2F%2Fwww.5paisa.com%2Fmutual-funds%2F']
for i in range(3):
    wd = webdriver.Chrome('C:/Users/C207641/Downloads/chromedriver')
    for j in range(len(urls)):
        mobile_url = urls[j] + "&form_factor=mobile"
        wd.get(mobile_url)
        sheet['A'+str(j+1)].value = urls[j]
        mobile_score = ''
        while mobile_score == '':
            try:
                mobile_score = WebDriverWait(wd, 120).until(EC.presence_of_element_located((By.XPATH, '//*[@id="performance"]/div[1]/div[1]/div[1]/a/div[2]'))).text
                mobile_score = int(mobile_score)
            except:
                continue
        if str(sheet['B'+str(j+1)].value) == 'None':
            sheet['B'+str(j+1)].value = mobile_score
        elif sheet['B'+str(j+1)].value < mobile_score:
            sheet['B'+str(j+1)].value = mobile_score
            print(mobile_url,mobile_score,'Updated')

        desktop_url = urls[j] + "&form_factor=desktop"
        desktop_score = ''
        while desktop_score == '':
            wd.get(desktop_url)
            try:
                desktop_score = WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="performance"]/div[1]/div[1]/div[1]/a/div[2]'))).text
                desktop_score = int(desktop_score)
            except:
                continue
        if str(sheet['C'+str(j+1)].value) == 'None':
            sheet['C'+str(j+1)].value = desktop_score
            print(urls[j],mobile_score,desktop_score)
        elif sheet['C'+str(j+1)].value < desktop_score:
            sheet['C'+str(j+1)].value = desktop_score
            print(desktop_url,desktop_score,'Updated')
    
    wd.close()

book.save("C:/Users/C207641/Downloads/pagespeed_scores.xlsx")
book.close()