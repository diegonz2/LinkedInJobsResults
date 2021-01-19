import pandas as pd
import plotly
import plotly.graph_objects as go
from datetime import date
from selenium import webdriver
from secrets import pw
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
from time import sleep


# Authentication and scraping of the number of jobs in a Linkedin search.
# Put LinkedIn username 'user@gmail.com'
LinkedInUsername = ""

chrome_options = Options()
chrome_options.add_argument("--headless")           # This is done to NOT display a Google Chrome window in our monitor while the data is being collected.
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://www.linkedin.com/")
driver.find_element_by_xpath("/html/body/main/section[1]/div[2]/form/div[2]/div[1]/input")\
            .send_keys(LinkedInUsername)
# Put password in the document 'secrets.py'
driver.find_element_by_xpath("/html/body/main/section[1]/div[2]/form/div[2]/div[2]/input")\
            .send_keys(pw)
driver.find_element_by_xpath("/html/body/main/section[1]/div[2]/form/button")\
            .click()
sleep(6)

# Put the URL of the Linkedin Jobs Search you want. Example (Keyworkd=Cloud , Location=Worldwide)
driver.get("https://www.linkedin.com/jobs/search/?geoId=92000000&keywords=cloud&location=Worldwide")
sleep(4)

# Get the number of jobs as a String
JobsStringText = driver.find_element_by_xpath("/html/body/div[7]/div[3]/div[3]/div/div/section[1]/div/header/div[1]/small").text

# Try to get only the number of jobs without a space at the end (example: '450,345 ')
number = ""
for i in JobsStringText:
    if i != " ":
        number = number + i
    else:
        break


# Replace the comma (',') in the number of jobs with an empty value ('')
JobsResults = int(number.replace(',', ''))

# Get the date when the script is run
today = date.today()


# Define the rows and columns to use in the Excel document
df = pd.DataFrame({'Jobs results': [JobsResults],
                   'Date': [today]})
writer = pd.ExcelWriter('Jobsresults.xlsx', engine='openpyxl')
# Try to open an existing workbook
writer.book = load_workbook('Jobsresults.xlsx')
# Copy existing sheets
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
# Read existing file
reader = pd.read_excel(r'Jobsresults.xlsx')
# Write out the new sheet
df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)
writer.close()


# Read the data written in the new excel column
excel_file = 'Jobsresults.xlsx'
df = pd.read_excel(excel_file)
print(df)

# Put the data in a graphic html file
data = [go.Scatter( x=df['Date'], y=df['Jobs results'])]
fig = go.Figure(data)
fig.update_layout(title_text="Linkedin 'Example'(Keyword) Jobs" , title_font_size=40 , title_x=0.5)
plotly.offline.plot(fig, filename="JobsResultsGraph.html")
