import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time


# Function to set up the WebDriver
def setup_driver(driver_path):
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service)
    return driver

# Function to read data from Excel
def read_data_from_excel(excel_path):
    return pd.read_excel(excel_path)


# Function to extract table data and check for company name match
def extract_table_data(driver, trade_name):
    data = {}
    company_name_status = "False"  # Default status
    rows = driver.find_elements(By.XPATH, "//div[@class='card-body']//table//tr")
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) == 4:
            key = cells[0].text[:-1]  # Remove colon at the end if present
            data[key] = cells[1].text
            # Check if one of the keys is 'Company Name' and if it matches the trade name
            if (
                key == "Company Name"
                and cells[1].text.strip().lower() == trade_name.strip().lower()
            ):
                company_name_status = "True"
            data[key] = cells[2].text
        elif len(cells) == 2:
            key = cells[0].text[:-1]
            data[key] = cells[1].text
            if (
                key == "Company Name"
                and cells[1].text.strip().lower() == trade_name.strip().lower()
            ):
                company_name_status = "True"
    return data, company_name_status

# Function to process emirates inquiry
def process_emirates_inquiry(driver, row, emirate_options):
    trade_name = row['trade_name_en']
    driver.get("https://inquiry.mohre.gov.ae/")

    data_found = False
    result_data = {"Status": "False", "Checked Emirates": ""}

    for emirate in emirate_options:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "js-example-basic-single"))
        )
        service_select = Select(driver.find_element(By.ID, "js-example-basic-single"))
        service_select.select_by_visible_text("Company Information by License No")

        transaction_input = driver.find_element(By.ID, "inputTransaction")
        transaction_input.clear()
        transaction_input.send_keys(row["license_number"])

        WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID, "jsEmirizationddlbasicsingle"))
        )
        Select(driver.find_element(By.ID, "jsEmirizationddlbasicsingle")).select_by_value(emirate["value"])

        captcha_span = WebDriverWait(driver, 2).until(
            EC.visibility_of_element_located((By.ID, "spanDisplayOtp"))
        )
        captcha = captcha_span.text

        captcha_input = driver.find_element(By.ID, "InputOTP")
        captcha_input.send_keys(captcha)

        search_button = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.ID, "btnSearchServiceData"))
        )
        search_button.click()

        try:
            WebDriverWait(driver, 2).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "card-body"))
            )
            extracted_data, company_name_status = extract_table_data(driver, trade_name)
            extracted_data["Status"] = "True"
            extracted_data["CompanyNameStatus"] = company_name_status
            extracted_data["Checked Emirates"] = emirate["name"]
            data_found = True
            result_data = extracted_data
            break
        except TimeoutException:
            continue

    if not data_found:
        result_data["Status"] = "No data found in any Emirate"
    
    return result_data

def extract_member_info(driver):
    member_info = []
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "gbox_s_2_l")))
    rows = driver.find_elements(By.XPATH, "//table[@id='s_2_l']/tbody/tr[@role='row' and contains(@class,'jqgrow')]")
    for row in rows:
        row_data = {}
        cells = row.find_elements(By.XPATH, ".//td")
        row_data['Member Name'] = cells[1].text
        row_data['Email'] = cells[2].text
        row_data['Phone'] = cells[3].text
        row_data['Fax'] = cells[4].text
        row_data['URL'] = cells[5].text
        row_data['Product/Service'] = cells[6].text
        row_data['Product Description'] = cells[7].text
        member_info.append(row_data)
    return member_info

def search_dubai_chamber(driver, trade_name):
    dubai_chamber_url = "https://eservice.dubaichamber.com/siebel/app/pseservice/enu/?SWECmd=GotoView&SWEBHWND=&_tid=1712058504&SWEView=DC+Commercial+Directory+Landing+View&SWEHo=eservice.dubaichamber.com&SWETS=1712058504"
    
    # Navigate to the Dubai Chamber website
    driver.get(dubai_chamber_url)
    
    # Wait for the input field to be ready and enter the trade name
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "s_1_1_6_0")))
    input_field = driver.find_element(By.NAME, "s_1_1_6_0")
    input_field.clear()
    input_field.send_keys(trade_name)

    # Now, click the Search button
    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "s_1_1_10_0_Ctrl"))
    )
    search_button.click()
    

    # Wait for the results or no results notification to be visible
    try:
        # Adjust the wait below as per the expected time for results to show up
        WebDriverWait(driver, 10).until(
            lambda d: d.find_element(By.ID, "gbox_s_2_l").is_displayed() or
                      d.find_element(By.ID, "no_result_element_id").is_displayed()
        )

        # If results are found, extract them
        if driver.find_elements(By.ID, "result_element_id"):
            return extract_member_info(driver)
        else:
            return {"Status": "No data found on Dubai Chamber website"}

    except TimeoutException:
        # Handle the case where neither results nor no-results notification is found within the timeout
        return {"Status": "Failed to retrieve data from Dubai Chamber website"}

    # Extract member information
    return extract_member_info(driver)


# Function to save results to an Excel file
def save_results_to_excel(df, file_path):
    df.to_excel(file_path, index=False)






def main():
    driver_path = r"E:\Captcha Bypass\chrome_driver\chromedriver.exe"
    excel_path = r"E:\Captcha Bypass\excel\sample.xlsx"
    output_path = r"E:\Captcha Bypass\excel\final_updated_sample.xlsx"

    driver = setup_driver(driver_path)
    data = read_data_from_excel(excel_path)

    # Define emirate options based on the website you will be scraping

    emirate_options = [
        {"name": "Abu Dhabi", "value": "000000001"},
        {"name": "Dubai", "value": "000000002"},
        {"name": "Sharjah", "value": "000000003"},
        {"name": "Ras Al Khaima", "value": "000000004"},
        {"name": "Ajman", "value": "000000005"},
        {"name": "Fujairah", "value": "000000006"},
        {"name": "Um Al Qaiwain", "value": "000000007"},
        {"name": "Zones-Corp", "value": "000000008"},
    ]

    # This DataFrame will store all results
    all_results_df = pd.DataFrame()

    try:
        for index, row in data.iterrows():
            # Process MOHRE inquiry for the current row
            mohre_results = process_emirates_inquiry(driver, row, emirate_options)
            # Process Dubai Chamber search for the current row
            dubai_chamber_results = search_dubai_chamber(driver, row['trade_name_en'])
            # Combine results into a single row DataFrame
            combined_results = {**mohre_results, **dubai_chamber_results}
            combined_results_df = pd.DataFrame([combined_results])
            # Append the combined results to the all_results_df DataFrame
            all_results_df = pd.concat([all_results_df, combined_results_df], ignore_index=True)
        
        # Save the combined results to an Excel file
        save_results_to_excel(all_results_df, output_path)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
