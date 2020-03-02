from cerberus import Validator
import re, datetime, csv, docx, os
#CHANGE LOG 2/20/19
# -Added data validation so users can not enter invalid data. This requires the cerberus library. Validation parameters
#  can be nmodified in the mainSchema. Use the getvalidData function to prompt user and make sure data is valid against supplied schema key.
# -Added a yesno function for all yes no questions needed of the user. This helps resuce redudent code and allows is to use booleans instead of string
# comparisons which is overall a better thing to be comparing and reduces code for if statements.
# -Added ability to change directory for letter and log in global variables below
# -Added error handeling to both functions that write to files. This is needed in case of a permission to write or the user has the file open.
# -Determining case section no longer askes for mercPOS if the chip desicion is a yes. Also does not ask for case number either
# -Added setting for asking for letter or not

'''

NO CHARGEBACK  == no claim

MCC=5968 == NO CHARGEBACK
Member chip=N & merchant chip=Y  == NO CHARGEBACK
Member chip=Y & merchant chip=Y/N == Claim




'''



#Paths to save log and letters
SAVE_PATH_ROOT = 'Desktop' #anything in users folder I believe
LETTER_FOLDER = '\\RegE Log 2019\Reg E Letters - Not Printed\\'
LOG_FOLDER = '\\RegE Log 2019\\'
LOG_NAME = 'Reg E Log 2019.csv'

#Ask for letter or not
noask = True

#This schema is used for validating user input using the cerberus libaray
#Check out the cerberus docs for information on validation rules and creating schemas
#http://docs.python-cerberus.org/en/stable/validation-rules.html

mainSchema={'name': {'maxlength': 15, 'empty': False, 'minlength': 4, 'regex': '^[a-zA-Z]+$'},
		'transAmount': {'regex': '[+-]?([0-9]*[.])?[0-9]+', 'empty': False}, 
		'address': {'empty': False, 'minlength': 8},
		'merchant': {'maxlength': 15, 'empty': False, 'minlength': 3, 'regex': '^[a-zA-Z]+$'},

		'number': {'regex': '[+-]?([0-9]*[.])?[0-9]+', 'empty': False, 'minlength': 1}, #Generic number validation floats and integer
		'string': {'regex': '^[a-zA-Z]+$', 'empty': False, 'minlength': 1}, #Generic string validation
		'yesno': {'regex': '^[yYYes]|^[nNNo]|^$', 'empty': False, 'minlength': 1} #Yes(Y/y) / No(N/n) validation
		} 



#This function takes in user input and validates it. You must provide the prompt for the user and the name in the schema to validate it agains.
#Optionally you can provide an error message for the user in case of invalid entered data
#getValidInput(Prompt for user, Name of rule in schema to validate against, optional error message, optional different schema)

def getValidInput(prompt, schemaName, errorMsg = 'Invalid Data Entered - Try Again!', schema=mainSchema):
	v = Validator()
	while True:
		data = input(prompt)
		if v.validate({schemaName:data},schema):
			return data
		else:
			print(errorMsg)

#Turn yes/no reponces into boolean True=yes, False=no
def yesno(dat):
	if re.match('^[yYYes]+$', dat) is not None:
		return True
	else:
		return False


#Create Log file
def writeLog():
	userpath = os.path.join(os.path.join(os.environ['USERPROFILE']), SAVE_PATH_ROOT)

	#Make folder if it does not exist
	if not os.path.exists(userpath + LOG_FOLDER):
		os.makedirs(userpath + LOG_FOLDER)


	with open(userpath+LOG_FOLDER+LOG_NAME, 'a', newline='') as f:

		merchantdelta = datetime.date.today() + datetime.timedelta(days=45)
		caseclosedelta = datetime.date.today() + datetime.timedelta(days=90)
		rowheaders = ['Date','Name','Address','Merchant','Amount', 'Claim Number', 'Case Number', 'Merchant Response Date', 'Case Closed Date']
		thewriter = csv.DictWriter(f, fieldnames=rowheaders)
		thewriter.writerow({'Date':datetime.date.today(),'Name':name,'Address':address,'Merchant':merchant,'Amount':transAmount,
		                    'Claim Number':claimnum,'Case Number':casenum, 'Merchant Response Date':merchantdelta,'Case Closed Date':caseclosedelta})
		print('The Reg E log has been updated. Double check everything and print the member a letter.')
		return Tru
#Write letters
def createworddoc(): 
	userpath = os.path.join(os.path.join(os.environ['USERPROFILE']), SAVE_PATH_ROOT)

	#Make folder if it does not exist
	if not os.path.exists(userpath + LETTER_FOLDER):
		os.makedirs(userpath + LETTER_FOLDER)

	doc = docx.Document()
	doc.add_paragraph(name, 'Normal')
	doc.add_paragraph(address, 'Normal')
	doc.add_paragraph('Thank you for reaching out to this company. '
	                  'we would like to offer you a discount on the ' + address + 'product you bought')
	doc.save(userpath + "\\RegE Log 2019\Reg E Letters - Not Printed\\" + name + '.docx')
	print("Letter Succesfully Written! Located here:", userpath + LETTER_FOLDER + name + '.docx')
	return True


def finalize():
	#Try creating log file
	while True:
		try:
			if writeLog():	
				break

		except:
			errMenu = getValidInput("Could not write to log! Please check that log is not open! Type Yes to try writing again and No to exit. :", 'string')
			if not yesno(errMenu):
				#No - exit
				break



	#Try writing letter
	askedOnce = False
	while True:
		try:
			if not askedOnce | noask:
				if yesno(getValidInput("Would you like to generate letter to client? (Y/N): ", 'yesno', 'Only Yes/Y/y or No/N/n accepted!')) | noask:
					#Yes generate letter
					if createworddoc():
						return
				else:
					return
			else:
				#Asked if user wanted to generate letter trying again from error
				if createworddoc():
						return

		except:
			errMenu = getValidInput("Could not write letter! Type Yes to try writing again and No to exit. :", 'string')
			askedOnce = True
			if not yesno(errMenu):
				#No - exit
				return



def merchCode():
	#Merchant Code
	mcc = getValidInput('What is the merchant category code: ','number', 'Must be a number!')
	if mcc == '5968':
		return True
	elif mcc != '5968':
		return False
		#print('Research if the merchant had a POS system that read chips and if the member card was a chip card')


#Chip decision 
def memberChip():
	chipdecision = getValidInput('Was the debit card used chip enabled? (Y/N) : ', 'yesno', 'Only Yes/Y/y or No/N/n accepted!')
	if yesno(chipdecision):
		#yes
		print('Thank you.')
		return True #Advantage
	else:
		#No
		print('Thank you.')
		return False #Disadvantage


#Merchant POS
def merchPOS():
	while True:
	    tracknumber = getValidInput('What was the merchant POS track number? : ', 'number')
	    if tracknumber == '101':
	        print('Thank you.')
	        return True #Advantage
	    elif tracknumber == '102':
	        print('Thank you.')
	        return False #Disatvantage
	    else:
	        print('The only known track numbers at this time are 101 and 102.')



#Gather client data
print('Hello. Welcome to the Reg D software.')
name = getValidInput('Enter the member name: ', 'name', 'Invalid Data! Name must not contain number! Length must be between 4-15 characters!')
address = getValidInput('Enter the members address: ', 'address', 'Invalid Data! Length must be at least 8 characters!')
transAmount = getValidInput('Enter the transaction amount:$', 'transAmount', 'Invalid Data! Numbers Only!')
merchant = getValidInput('Enter the merchant: ', 'merchant', 'Invalid Data! Name must not contain number! Length must be between 3-15 characters!')


#Determine case submission 
casenum ="N/A"
claimnum = "N/A"
if merchCode():
	#Was gas pump NO CHARGEBACK
	print('This transaction was a gas pump.  Submit the case as Fraud Advice instead of Enhanced Dispute Request, no chargeback rights.')
	finalize()

elif not memberChip() and merchPOS():
	#Member had no chip and merchant can read chip NO CHARGEBACK
	print('The merchant was capable of reading chip cards but the members debit card was not chip enabled.  Submit the case as Fraud Advice instead of Enhanced Dispute')
	finalize()

else:
	#Member had chip 
	print('Submit the transaction with Fiserv.')
	#Get case and claim number
	casenum = getValidInput('What is the case number of the submitted chargeback? :', 'number')
	claimnum = getValidInput('What is the claim number of the submitted chargeback? :', 'number')
	finalize()





















	