import pandas as pd  # Import the pandas library for data manipulation
import os  # Import the os module for interacting with the operating system
import smtplib  # Import the smtplib module for sending emails
import ssl  # Import the ssl module for secure connections

from email.message import EmailMessage  # Import the EmailMessage class for creating email messages
from selenium import webdriver  # Import Selenium webdriver for web scraping
from selenium.webdriver.common.by import By  # Import the By class for locating elements using different strategies
from selenium.webdriver.support.ui import WebDriverWait  # Import WebDriverWait for waiting for specific conditions
from selenium.webdriver.support import expected_conditions as EC  # Import expected_conditions for defining expected conditions
from selenium.common import NoSuchElementException  # Import the NoSuchElementException class for handling missing elements


# Define a function to extract data
def extract_data():
    offices_and_addresses = driver.find_elements(By.CLASS_NAME, "margin-16")  # Find elements with the specified class
    for curr_office in offices_and_addresses:
        name = curr_office.find_element(By.XPATH, './p[1]').text  # Find and extract text from the first paragraph under the current office
        address = curr_office.find_element(By.XPATH, './p[2]').text  # Find and extract text from the second paragraph under the current office
        phone_number = curr_office.find_element(By.XPATH, './dl/dd/p').text  # Find and extract text from a specific paragraph under the current office
        try:
            sat_hours = curr_office.find_element(By.XPATH, './div[1]/div/dl/dd/span/p[2]').text  # Try to find and extract text for Saturday working hours
        except NoSuchElementException:
            sat_hours = ""  # If NoSuchElementException occurs, set Saturday working hours to an empty string
        try:
            sun_hours = curr_office.find_element(By.XPATH, './div[1]/div/dl/dd/span/p[3]').text  # Try to find and extract text for Sunday working hours
        except NoSuchElementException:
            sun_hours = ""  # If NoSuchElementException occurs, set Sunday working hours to an empty string

        # Add the extracted data to the respective lists
        office_name.append(name)
        office_addresses.append(address)
        phone_numbers.append(phone_number)
        saturday.append(sat_hours)
        sunday.append(sun_hours)


# Define a function to export data to an Excel file
def export():
    df = pd.DataFrame({
        "Име на офиса": office_name,
        "Адрес": office_addresses,
        "Телефон": phone_numbers,
        "Раб.време събота": saturday,
        "Раб.време неделя": sunday
    })
    try:
        os.remove("C:\\PythonApp\\fibank_branches.xlsx")  # If the file already exists in the directory, then remove it
    except FileNotFoundError:  # If FileNotFoundError occurs, then skip
        pass
    finally:
        df.to_excel('C:\\PythonApp\\fibank_branches.xlsx', index=False)  # Export the DataFrame to an Excel file


# Define a function to send an email with attached data
def send_email():
    email_sender = "lflorov2@gmail.com"
    email_password = "fcuo esuy qbeq vaqe"
    email_receiver = "db.rpa@fibank.bg"

    host = "smtp.gmail.com"
    port = 465

    subject = "Информация за офисите на Fibank"
    body = """
Здравейте,
Прикачен е файл с информация за офисите на Fibank работещи събота и неделя.

Поздрави,
Любослав Флоров
"""

    # Set the name and path for the file
    file_name = "fibank_branches.xlsx"
    file_path = os.path.join("C:", "PythonApp", file_name)

    em = EmailMessage()  # Create an object of the EmailMessage class
    em["From"] = email_sender
    em["To"] = email_receiver
    em["Subject"] = subject
    em.set_content(body)  # Set the content of the email

    # Attach the Excel file to the email
    with open(file_path, "rb") as attachment:
        em.add_attachment(attachment.read(), maintype='application', subtype='octet-stream', filename=file_name)

    context = ssl.create_default_context()

    # Connect to the SMTP server, login, and send the email
    with smtplib.SMTP_SSL(host, port, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())


# Set the URL of the website
website = "https://my.fibank.bg/EBank/public/offices"
driver = webdriver.Chrome()  # Create an instance of the Chrome WebDriver
driver.get(website)  # Open the website in the browser

# Initialize empty lists to store the extracted data
office_name = []
office_addresses = []
phone_numbers = []
saturday = []
sunday = []

# Create an instance of WebDriverWait for waiting for specific conditions
wait = WebDriverWait(driver, 10)

# Wait for the button to be clickable before clicking
button_xpath = '//*[@id="content-col"]/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/button'
wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath))).click()

# Wait for the link to be clickable before clicking
link_xpath = ('//*[@id="content-col"]/div[2]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/ul/li['
              '3]/a/span[1]')
wait.until(EC.element_to_be_clickable((By.XPATH, link_xpath))).click()

# Extract data using the defined function
extract_data()

# Close the browser
driver.quit()

# Export and send the email
export()
send_email()
