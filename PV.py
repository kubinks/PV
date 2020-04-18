import requests
import time
from selenium import webdriver
import pyautogui
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import Tk
from io import StringIO
import sqlite3
from gspread_formatting import *
import string

def sleep(t):
    time.sleep(t)

year_now = time.strftime('%Y', time.localtime())
month_now = int(time.strftime('%m', time.localtime()))

########################### FRONIUS ###########################
#opening chrome
driver = webdriver.Chrome()
#going to inverter's login page
driver.get('https://www.solarweb.com/Account/ExternalLogin')
#logging in
id_box = driver.find_element_by_id('username')
id_box.send_keys('---------------')

pass_box = driver.find_element_by_id('password')
pass_box.send_keys('---------------')
sleep(2)

login_button = driver.find_element_by_id('submitButton')
login_button.click()
sleep(3)
#accepting cookies
cookie = driver.find_element_by_xpath('//*[@id="cookiebanner"]/div/div/p[2]/a[@id="CybotCookiebotDialogBodyButtonAccept"]')
cookie.click()
sleep(3)
#navigating to tab with data
driver.get('https://www.solarweb.com/Chart/Chart?pvSystemId=223a1bcc-1cf6-486e-915f-e569e0e58b58')
sleep(3)

#opening devtools and navigating through it on my 2560x1080 monitor
#devtools
pyautogui.keyDown('ctrl')
pyautogui.keyDown('shift')
pyautogui.press('j')
pyautogui.keyUp('ctrl')
pyautogui.keyUp('shift')
sleep(1)
#network
pyautogui.moveTo(1025, 149)
pyautogui.click()
sleep(1)
#XHR
pyautogui.moveTo(756, 224)
pyautogui.click()
sleep(1)
#refresh
pyautogui.keyDown('ctrl')
pyautogui.press('r')
pyautogui.keyUp('ctrl')
sleep(3)
#Rok
pyautogui.moveTo(389, 901)
pyautogui.click()
sleep(1)
#GetChartNew
pyautogui.moveTo(816, 419)
pyautogui.click()
sleep(1)
#settings
pyautogui.moveTo(987, 405)
pyautogui.click()
sleep(1)
#series
pyautogui.moveTo(976, 420)
pyautogui.click()
sleep(1)
#zero
pyautogui.moveTo(970, 436)
pyautogui.click()
sleep(1)
#data
pyautogui.moveTo(1083, 482)
sleep(1)
#triple click + copy
pyautogui.click()
pyautogui.click()
pyautogui.click()
pyautogui.keyDown('ctrl')
pyautogui.press('c')
pyautogui.keyUp('ctrl')
sleep(1)
#saved copied text to a variable
clipboard = Tk().clipboard_get()
#close chrome
pyautogui.keyDown('altleft')
pyautogui.keyDown('f4')
pyautogui.keyUp('altleft')
pyautogui.keyUp('f4')

#reading file
# fronius = pd.read_csv('c:/Users/KubaPC/Desktop/fronius.txt')
fronius = pd.read_csv(StringIO(clipboard), sep=',', header=None)
#creating a list of produciton per month
fronius = fronius.values.tolist()[0]
fronius_data = []
for i,item in enumerate(fronius):
    #we know that our data is stored in odd positions
    if i%2!=0:
        fronius_data.append(item.replace(']', '').replace('.', ','))

#connecting to google sheets
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('employees_secret.json', scope)
client = gspread.authorize(creds)

#choosing spreadsheet and worksheet
sh = client.open('BILANS PV N')
worksheet = sh.get_worksheet(3)

#getting values from column with years to have the reference
values_list = worksheet.col_values(4)
months = {}
#filling the dictionary with keys as years and values as lists of rows where each month is stored
for i,item in enumerate(values_list):
    if item == '':
        continue
    else:
        if item in months:
            months[item].append(i+1)
        else:
            months[item] = [i+1]

#making a dictionary to translate column numbers into letters
d = dict(enumerate(string.ascii_uppercase, 1))
#inserting data to worksheet
for i in range(len(fronius_data)):
    worksheet.update_cell(months[year_now][0]+i, 16, fronius_data[i])
    print(f'updating PV, col {d[16]}, row {months[year_now][0]+i} with value {fronius_data[i]}')

########################### ONEMETER ###########################

#connecting to onemeter
APIkey = '---------------'
device_id = '---------------'
device_id2 = '---------------'
url_usages = f'https://cloud.onemeter.com/api/devices/{device_id}/export/day/data.xlsx'
url_usages2 = f'https://cloud.onemeter.com/api/devices/{device_id2}/export/day/data.xlsx'
usages = requests.get(url_usages, headers={'Authorization': APIkey}, params={'columns': ['1_8_0', '2_8_0']})
usages2 = requests.get(url_usages2, headers={'Authorization': APIkey}, params={'columns': ['1_8_0', '2_8_0']})
#loading downloaded data into dataframes and merging them
cols = ['Data', '1.8.0', '2.8.0']
df = pd.read_excel(usages.content, skiprows=4, header=None)
df2 = pd.read_excel(usages2.content, skiprows=4, header=None)
df.columns = cols
df2.columns = cols
df = pd.concat([df2, df])

fltr_m = df.Data.dt.month
fltr_y = df.Data.dt.year
#going through every year and month in downloaded data
for yr in sorted(fltr_y.unique().tolist()):
    for mth in sorted(fltr_m.unique().tolist()):
        month = df[(fltr_m == mth) & (fltr_y == yr)]
        if mth == 12:
            month_next = df[(fltr_m == 1) & (fltr_y == yr+1)]
        else:
            month_next = df[(fltr_m == mth+1) & (fltr_y == yr)]
        #testing to not get an error by substracting from next month when its empty
        try:
            result = month_next.iloc[0] - month.iloc[0]
            try:
                cons_mth = months[str(yr)][mth-1]
                taken = str(round(result[1], 2)).replace('.', ',')
                sent = str(round(result[2], 2)).replace('.', ',')
                worksheet.update_cell(cons_mth, 9, taken)
                worksheet.update_cell(cons_mth, 11, sent)
                print(f'updating PV, col {d[9]}, row {cons_mth} with value {taken}')
                print(f'updating PV, col {d[11]}, row {cons_mth} with value {sent}')
                #print(f'Year: {yr}, month: {mth}, taken: {taken}, sent: {sent}')
            except:
                continue
        except:
            continue

########################### DOMOTICZ ###########################
    
#setting up the dictionary of devices with their IDX
devices = {'1 - biuro': 33, '2 - taśmowa': 34, '3 - wejście': 32,
              '4 - zlew': 35, '5 - prasa ogród': 29, '6 - prasa ulica': 37,
              'Malarnia': 38, 'Lustro': 39, 'Klima': 85}
#connecting to the downloaded database
dz = sqlite3.connect('C:/Users/KubaPC/Downloads/Domoticz.db')
c = dz.cursor()

#particular months are in the same rows hence no need to generate months dict
#changing worksheet to 'Xiaomi'
worksheet = sh.get_worksheet(4)
row_values = worksheet.row_values(3)
#coloring input data due to daily approximation in domoticz database
fmt = cellFormat(textFormat=textFormat(foregroundColor=color(1, 0, 0)))
#iteration for matching the device with its column
for i, name in enumerate(row_values):
    #print(f'matching {name} with devices')
    for key in devices.keys():
        if key == name:
            #print(f'match found on col {d[i+1]}')
            #after match is found, querying db for data
            for row in c.execute(f'''SELECT strftime('%m', Date), SUM(Value)
                    FROM Meter_Calendar
                    WHERE DeviceRowID = {devices[key]} AND strftime('%Y', Date) = '{year_now}'
                    GROUP BY strftime('%m', Date)
                    ORDER BY Date ASC'''):
                cons_mth = months[year_now][int(row[0])]-1
                #print('cell ', worksheet.cell(cons_mth+1, i+1).value)
                #checking for 'x' meaning that this row has been validated
                if worksheet.cell(cons_mth, 20).value != 'x':
                    worksheet.update_cell(cons_mth, i+1, row[1]/1000)
                    format_cell_range(worksheet, 'E{0}:M{0}'.format(cons_mth), fmt)
                    print(f'updating Xiaomi, col {key}, row {cons_mth} with value {row[1]/1000}')
dz.close()
##TODO: download domoticz.db
##TODO: make a list for particular year and update whole year at once