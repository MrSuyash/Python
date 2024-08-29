import json
import re
import xlsxwriter

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

# Initialize the Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
url = "https://www.imdb.com/chart/top"
driver.get(url)

# Wait for the page to load completely and elements to be present
WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='ipc-title-link-wrapper']/h3[@class='ipc-title__text']")))

# XPaths to extract data
title_xpath = "//div[@class='ipc-title ipc-title--base ipc-title--title ipc-title-link-no-icon ipc-title--on-textPrimary sc-b189961a-9 bnSrml cli-title']/a[@class='ipc-title-link-wrapper']/h3[@class='ipc-title__text']"
year_xpath = "//div[@class='sc-b189961a-7 btCcOY cli-title-metadata']/span[1]"
rating_xpath = "//span[@class='ipc-rating-star--rating']"
num_rating_xpath="//span[@class='ipc-rating-star--voteCount']"

# Extract the movie titles
titles = driver.find_elements(By.XPATH, title_xpath)
# Extract the release years
years = driver.find_elements(By.XPATH, year_xpath)
# Extract the IMDb ratings
ratings = driver.find_elements(By.XPATH, rating_xpath)
#Extract the IMDB no. of ratings
nrating = driver.find_elements(By.XPATH, num_rating_xpath)
# Prepare data for DataFrame
data = {
    "Title": [title.text for title in titles],
    "Year": [year.text for year in years],
    "IMDb Rating": [rating.text for rating in ratings],
    "No. Of Rating":[n.text for n in nrating]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
excel_filename = "imdb_top_250.xlsx"
df.to_excel(excel_filename, index=False)

# Close the WebDriver
driver.quit()
