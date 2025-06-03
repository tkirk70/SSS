import requests
import json
import base64
import configparser

config = configparser.ConfigParser()
config.read('config.ini')

client_id = config['Surface']['CLIENT_ID']
client_secret = config['Surface']['CLIENT_SECRET']
number_id = config['Surface']['NUMBER_ID']
customer_id = config['Surface']['CUSTOMER_ID']

url = "https://secure-wms.com/AuthServer/api/Token"

payload = json.dumps({
  "grant_type": "client_credentials",
  "user_login_id": number_id
})

headers = {
    'Host': 'secure-wms.com',
    'Connection': 'keep-alive',
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Authorization': f"Basic {base64.b64encode(bytes(f'{client_id}:{client_secret}', 'utf-8')).decode()}",
    'Accept-Encoding': 'gzip,deflate,sdch',
    'Accept-Language': 'en-US,en;q=0.8',
}
    
response = requests.request("POST", url, headers=headers, data=payload)

j = response.json()

url = f'https://secure-wms.com/inventory/stockdetails?customerid={customer_id}&facilityid=4&pgsiz=200'
payload={}
headers = {
  'Accept-Language': 'en-US,en;q=0.8',
  'Host': 'secure-wms.com',
  'Content-Type': 'application/json; charset=utf-8',
  'Accept': 'application/hal+json',
  'Authorization': 'Bearer '+j['access_token']
}

response = requests.request("GET", url, headers=headers, data=payload)

j = response.json()

# Create dictionary for df? Or create df then append it?

data = json.loads(response.text)

# for item in data['_embedded']['item']:
#     sku = item['itemIdentifier']['sku']
#     quantity = item['onHand']
#     desc = item['description']
    # print(f'SKU: {sku}, Description: {desc}, Quantity: {quantity}')

import pandas as pd
df = pd.DataFrame()

# append columns to an empty DataFrame
df['SKU'] = [item['itemIdentifier']['sku'] for item in data['_embedded']['item']]
  
df['DESC'] = [item['description'] for item in data['_embedded']['item']]

# df['DESC2'] = [item['description2'] for item in data['_embedded']['item']]

# df['EXP'] = [item['expirationDate'] for item in data['_embedded']['item']]
df['EXP'] = [item.get('expirationDate') for item in data.get('_embedded', {}).get('item', [])]
df['EXP'] = pd.to_datetime(df['EXP']).dt.strftime('%Y-%m-%d')
  
df['QTY'] = [int(item['available']) for item in data['_embedded']['item']]



# df['LOC'] = [item.get('locationIdentifier', {}).get('nameKey', {}).get('name', 'Unassigned') for item in data['_embedded']['item']]

# df['LOC'] = [item['locationIdentifier']['nameKey']['name'] if 'locationIdentifier' in item and 'nameKey' in item['locationIdentifier'] and 'name' in item['locationIdentifier']['nameKey'] else 'Unassigned' for item in data['_embedded']['item']]

# # Set to integer and format for thousands seperator - this messed something up.
# df['QTY'] = df['QTY'].astype(int)
# df['QTY'] = df['QTY'].map('{:,}'.format)

# Pull in yesterday's report to compare
# df_prev = pd.read_excel('Previous Lux Row Inventory.xlsx')
# df_prev.index = df_prev.index + 1
# The above works.  Create a blank df and create column headers and append with comprehension list.

import time

# Get today's date.
from datetime import date
today = date.today()

time.sleep(5)

# This works for Lux Row as well.

# How often to send, what format and to whom?

# Do I generate a df or a simple email?

# record changes day over day?

# Set max rows
pd.set_option('display.max_rows', None)

# Sort by sku alphabetical

df.sort_values(by='SKU', ascending=True)

# Combine same SKU's and show aggregate

df = df.groupby(['SKU', 'DESC']).sum().reset_index()

df.index = df.index + 1

sku_list = [
  
]

cost_list = [
 
]

cost_dict = dict(zip(sku_list, cost_list))

df['COST'] = df['SKU'].map(cost_dict)

df['TOTAL'] = df['QTY'] * df['COST']

# Formatting COST and TOTAL with two decimal places and a thousands separator
df['QTY'] = df['QTY'].apply(lambda x: f"{x:,.0f}")
df['COST'] = df['COST'].apply(lambda x: f"${x:,.2f}")
df['TOTAL'] = df['TOTAL'].apply(lambda x: f"${x:,.2f}")

# df['Change'] = df['QTY'] - df_prev['QTY']

df.to_excel(f'BCG Sunscreen LLC Inventory for {today}.xlsx')

time.sleep(5)

import win32com.client

outlook = win32com.client.Dispatch('outlook.application')

mail = outlook.CreateItem(0)

mail.To = 'oliver@surfacesunscreen.com'
mail.Subject = f'BCG Sunscreen LLC {today}'
mail.HTMLBody = f'''<h3>Please find inventory for {today} below and in attached excel sheet.</h3> {df.to_html()}'''
# mail.Body = f"Hello; all.  Here is the current inventory for Lux Row & Upper Right.\n\n{df}"
mail.Attachments.Add(f'C:/Users/TechnologyDataServic/OneDrive - TCG3PL/Documents/3PL Central/Custom API/BCG Sunscreen LLC Inventory for {today}.xlsx')
mail.CC = 'tds@tcg3pl.com'

# Resets the options
pd.reset_option('all')

mail.Display(True)
# mail.Send()

# Create df to compare to tomorrow's
# df.to_excel('Previous Lux Row Inventory.xlsx')




'''You can use the Windows Task Scheduler to run a Python script daily. Here are the steps:

Create a batch file that runs your Python script.
Open the Task Scheduler application on your Windows machine.
Click on ‘Create Basic Task…’ in the Actions Tab.
Give a suitable name and description of your task that you want to schedule.
Select at what time you want to run the script daily.
Choose to start the task ‘Daily’ since we wish to run the Python script daily at 6 am.
Specify the start date and time (6 am for our example).
Select ‘Start a program’, and then press Next.
Use the Browse button to find the batch file (run_python_script.bat) that runs the Python script.
Here is a link to an article that explains how to schedule a Python script using Windows Scheduler 1. You can also check out this article on GeeksforGeeks 2 for more information.

I hope this helps! Let me know if you have any other questions.'''


'''
There are several reasons why a batch file might not run properly in Task Scheduler. Here are some things you can try to troubleshoot the issue:

1. Check file/folder permissions: Ensure that the account you are using to run the script in Task Scheduler has Full Control permissions on the folder containing the script, the script itself, and any folders/files that the script touches when it runs ¹.
2. Specify the full path of the batch file: For.bat files to run inside your scheduled task, you need to specify your.bat file path inside the start option - despite the fact that your.bat file is at the same directory as your.exe ³.
3. Use cmd.exe to run the batch file: Make sure you run it using cmd.exe and add /c so that the cmd closes after the batch file has run ⁴.
4. Check for unsuitable characters in filename: Some characters are disliked by Task Scheduler. If your filename contains such characters, it may refuse to run ².

I hope this helps you resolve your issue with Task Scheduler! Let me know if you have any other questions.

Source: Conversation with Bing, 6/15/2023(1) Fix Scheduled Task Won’t Run for .BAT File - Help Desk Geek. https://helpdeskgeek.com/help-desk/fix-scheduled-task-wont-run-bat-file/ Accessed 6/15/2023.
(2) Batch runs manually but not in scheduled task - Stack Overflow. https://stackoverflow.com/questions/12513264/batch-runs-manually-but-not-in-scheduled-task Accessed 6/15/2023.
(3) .BAT file not running in task scheduler - Stack Overflow. https://stackoverflow.com/questions/30770042/bat-file-not-running-in-task-scheduler Accessed 6/15/2023.
(4) Windows Task Scheduler doesn't start batch file task. https://stackoverflow.com/questions/19318494/windows-task-scheduler-doesnt-start-batch-file-task Accessed 6/15/2023.
'''
