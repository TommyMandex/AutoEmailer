from selenium import webdriver

from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.common.keys import Keys


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
	
targetEmails = []
email = ""
password = ""
subject = ""
message = []

# Gets the target emails
with open("emails.txt", "r") as file:
	for line in file:
		# If the current line is not a comment, add the email to the list
		if not isComment(line):
			targetEmails.append(cleanLine(True, True, line))

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

capa = DesiredCapabilities.CHROME
capa["pageLoadStrategy"] = "none"

driver = webdriver.Chrome(desired_capabilities=capa)

driver.set_page_load_timeout(20)

# Login
driver.get("https://outlook.live.com/owa/?nlp=1")
wait = WebDriverWait(driver, 20)

# Wait for inital page to load
wait.until(EC.presence_of_element_located((By.NAME, 'loginfmt')))

driver.find_element_by_name("loginfmt").send_keys(email)
time.sleep(0.5)
driver.find_element_by_id("idSIButton9").send_keys(Keys.ENTER)
time.sleep(0.5)
driver.find_element_by_id("i0118").send_keys(password)
time.sleep(0.5)
driver.find_element_by_id("idSIButton9").send_keys(Keys.ENTER)

# Wait for webpage to load
time.sleep(5)


# Create a new email
for curEmail in targetEmails:
	# Click new message
	content = driver.find_elements_by_class_name("ms-Button--commandBar")
	content[0].send_keys(Keys.ENTER)
	time.sleep(2)
	# Adds sender address, the subject, and the message
	senderAddressInput = driver.find_element_by_class_name("pickerInput_8d9d7e4e")
	senderAddressInput.send_keys(curEmail)

	subjectInput = driver.find_element_by_class_name("field-150")
	subjectInput.send_keys(subject)
	bodyInput = driver.find_element_by_class_name("_2s9KmFMlfdGElivl0o-GZb")
	
	# Add body with a new line after each item in list
	for curLine in message:
		bodyInput.send_keys(curLine)
		bodyInput.send_keys(Keys.ENTER)
		time.sleep(0.5)
	
	# Send message
	sendButton = driver.find_element_by_class_name("y8XIN4EAeheOqsn4BJB7R")
	ActionChains(driver).move_to_element(sendButton).click(sendButton).perform()



# End Program
time.sleep(4)
driver.quit()
