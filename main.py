from selenium import webdriver

from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait

from selenium.webdriver.common.keys import Keys

from selenium.common.exceptions import TimeoutException
from datetime import datetime
import pathlib

import time

# Determines whether or not a text line is designated as a comment
def isComment(line):
	if line[0] == "/" and line[1] == "/":
		return True
	else:
		return False

# Removes Carriage Returns and spaces from around a string
def cleanLine(cleanSpace, cleanCarraigeReturn, line):
	# Removes leading and trailing whitespace
	if cleanSpace:
		line = line.strip()

	# Gets ride of carriage return
	if cleanCarraigeReturn and line.endswith("\n"):
		line = line.replace("\n", "")
	return line

version = "1.0.5v"
targetEmails = []
email = ""
password = ""
subject = ""
message = []
fileNames = []

print("Loading emails...")
# Gets the target emails
with open("emails.txt", "r") as file:
	for line in file:
		# If the current line is not a comment, add the email to the list
		if not isComment(line):
			targetEmails.append(cleanLine(True, True, line))
print("Successfully loaded " + str(len(targetEmails)) + " emails.")

print("Loading credentials...")
# Gets the login credentails
with open("credentials.txt", "r") as file:
	for line in file:
		# If the current line is not a comment
		if not isComment(line):
			# Get the content on right side of identifier and filter it
			parts = line.split(":")
			unfilteredLine = parts[1]
			filtered = cleanLine(True, True, unfilteredLine)
			# Determine if its the email or password
			if "email" in parts[0]:
				email = filtered
			elif "password" in parts[0]:
				password = filtered
print("Successfully loaded credentials.")

print("Loading email subject and message...")
# Gets the email subject and message
with open("message.txt", "r") as file:
	for line in file:
		if not isComment(line):
			# Check if this line is the subject line
			if line.startswith("subject:"):
				parts = line.split(":")
				unfilteredLine = parts[1]
				filtered = cleanLine(True, True, unfilteredLine)
				subject = filtered
			else:
				message.append(cleanLine(True, True, line))

print("Successfully loaded " + str(len(message)) + " message lines and loaded the subject '" + subject + "'")

print("Loading files to attach...")
# Gets the names of the files you want to upload
with open("files.txt", "r") as file:
	for line in file:
		if not isComment(line):
			fileNames.append(cleanLine(True, True, line))
print("Successfully loaded " + str(len(fileNames)) + " files to attach.")


print("Successfully loaded all files. Running version: '" + version + "'")
print("---------------------------------------------------")
print("Loading Chrome...")
capa = DesiredCapabilities.CHROME
capa["pageLoadStrategy"] = "none"

driver = webdriver.Chrome(desired_capabilities=capa)

driver.set_page_load_timeout(20)
print("Chrome successfully loaded.")
# Login
print("Fetching webpage...")
driver.get("https://outlook.live.com/owa/?nlp=1")
wait = WebDriverWait(driver, 20)
print("Webpage fetched.")

print("Starting log-in process")
# Wait for inital page to load
wait.until(EC.presence_of_element_located((By.NAME, 'loginfmt')))

driver.find_element_by_name("loginfmt").send_keys(email)
time.sleep(1)
driver.find_element_by_id("idSIButton9").send_keys(Keys.ENTER)
time.sleep(1)
driver.find_element_by_id("i0118").send_keys(password)
time.sleep(1)
driver.find_element_by_id("idSIButton9").send_keys(Keys.ENTER)
print("Log-in successful")

print("Pausing for webpage to load...")
# Wait for webpage to load
wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ms-Button--commandBar")))
time.sleep(5)
print("Pause complete.")
print("Begun sending emails.")
print("---------------------------------------------------")
log = open("log.txt", "a")
log.write("\n---------------------------------------------------")
curNum = 0
# Create a new email
for curEmail in targetEmails:
	curNum += 1
	print("Started sending email number: " + str(curNum))
	# Click new message
	content = driver.find_elements_by_class_name("ms-Button--commandBar")
	content[0].send_keys(Keys.ENTER)
	wait.until(EC.presence_of_element_located((By.CLASS_NAME, "pickerInput_8d9d7e4e")))
	time.sleep(1)
	# Adds sender address, the subject, and the message
	senderAddressInput = driver.find_element_by_class_name("pickerInput_8d9d7e4e")
	senderAddressInput.send_keys(curEmail)

	subjectInput = driver.find_element_by_class_name("ms-TextField-field")
	subjectInput.send_keys(subject)
	bodyInput = driver.find_element_by_class_name("_2s9KmFMlfdGElivl0o-GZb")
	
	# Add body with a new line after each item in list
	for curLine in message:
		bodyInput.send_keys(curLine)
		bodyInput.send_keys(Keys.ENTER)
		time.sleep(0.5)

	# Attaching files, if any
	attachmentFile = driver.find_element_by_class_name("_3V5k9IEYQL5bVRI9nrKNj-")
	for file_name in fileNames:
		attachmentFile.send_keys(str(pathlib.Path(__file__).parent.absolute()) + "/files_to_attach/" + file_name)

	# Waiting until the file is uploaded

	SHORT_TIMEOUT  = 5   # give enough time for the loading element to appear
	LONG_TIMEOUT = 30  # give enough time for loading to finish
	LOADING_ELEMENT_CLASS_NAME = 'ms-ProgressIndicator'
	if len(fileNames) > 0:
		print("Attaching files.")
		try:
			# wait for loading element to appear
			# - required to prevent prematurely checking if element
			#   has disappeared, before it has had a chance to appear
			WebDriverWait(driver, SHORT_TIMEOUT
				).until(EC.presence_of_element_located((By.CLASS_NAME, LOADING_ELEMENT_CLASS_NAME)))

			# then wait for the element to disappear
			WebDriverWait(driver, LONG_TIMEOUT
				).until_not(EC.presence_of_element_located((By.CLASS_NAME, LOADING_ELEMENT_CLASS_NAME)))

		except TimeoutException:
			pass 
		print("Files attached in email to '" + curEmail + "' .")
	else:
		print("No files attached to email.")
	

	
	
	# time.sleep(5)

	# Send message
	sendButton = driver.find_element_by_class_name("y8XIN4EAeheOqsn4BJB7R")
	ActionChains(driver).move_to_element(sendButton).click(sendButton).perform()

	# Logging time and date of email sent
	now = datetime.now()
	current_time = now.strftime("%m/%d/%Y, %H:%M:%S")
	print(current_time + ": Email number " + str(curNum) + " sent to '" + curEmail + "'")
	log.write("\n" +current_time + ": Email sent to '" + curEmail + "'")
	print("---------------------------------------------------")

print("Program completed. Sent " + str(curNum) + " emails.")

# End Program
time.sleep(4)
log.close()
driver.quit()
