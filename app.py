# Imports
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import concurrent.futures
from glob import glob
import xlsxwriter
import os
import time
import pandas as pd


# Path to chromedriver
PATH = r"C:\Program Files (x86)\chromedriver.exe"

# Initialize Selenuim Webdriver
# * Comment out when using multithreading
driver = webdriver.Chrome(PATH)
driver.get("http://35.247.68.174:8080/")

# Prepare data for testing
source_dirs = ["live", "voice"]
voice_calls = []
for dir in source_dirs:
    if dir == "live":
        voice_calls += glob(dir + '/*.mp3')
    elif dir == "voice":
        voice_calls += glob(dir + '/*.wav')

source_dirs = ["live(noise)", "voice(noise)"]
voice_calls_noise = []
for dir in source_dirs:
    if dir == "live(noise)":
        voice_calls += glob(dir + '/*.mp3')
    elif dir == "voice(noise)":
        voice_calls += glob(dir + '/*.wav')


# # Prediction list
# ('FILENAME', 'ACTUAL LABEL',
#                 'PREDICTED LABEL', 'WRONG FILE ANALYSIS', 'MORE FINDINGS')
predictions = [('FILENAME', 'ACTUAL LABEL',
                'PREDICTED LABEL(NOISE)', 'PREDICTED LABEL(NO NOISE)', 'WRONG FILE ANALYSIS', 'MORE FINDINGS')]

# Function def to check prediction


def check_prediction(filename):
    #! Uncomment these 2 lines when using multithreading
    # driver = webdriver.Chrome(PATH)
    # driver.get("http://35.247.68.174:8080/")

    # Get form tag
    form = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "form"))
    )
    # * print("Form done")

    # Get input field
    input_file = form.find_element_by_name("audio_file")
    # * print("Input got")

    # Upload File
    file = os.path.abspath(filename)
    input_file.send_keys(file)
    # * print("File uploaded")

    # Submit form
    submit_button = form.find_element_by_class_name("btn")
    submit_button.click()
    # * print("Submit clicked")

    # Get prediction
    prediction = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "pred-result"))
    )
    # * print("Pred got")

    # Get file extension for file
    # _, audio_file = filename.split("\\")
    _, file_extension = os.path.splitext(file)
    if file_extension == ".wav":
        actual_label = "Voice"
    elif file_extension == ".mp3":
        actual_label = "Live"

    # Update predictions
    predictions.append((filename, actual_label, prediction.text))
    # * print("List updated")

    #! Uncomment these 2 lines when using multithreading
    # print(file, " predicted!")
    # driver.quit()


# Run the function in for loop
# * Comment out when using multithreading
for file in voice_calls:
    check_prediction(file)
    print(file, " predicted!")

for file in voice_calls_noise:
    check_prediction(file)
    print(file, " predicted!")

#! Multithreading
# with concurrent.futures.ThreadPoolExecutor() as executor:
#     executor.map(check_prediction, voice_calls)

# Close chrome
# * Comment out when using multithreading
driver.quit()

# Make Excel file
with xlsxwriter.Workbook('testing without noise data.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    for row_num, data in enumerate(predictions):
        worksheet.write_row(row_num, 0, data)
