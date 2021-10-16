#############################################################################################################
#
# ┏┓ ┏┓ ┏━━━┓ ┏━━━┓ ┏━━━┓ ┏━━━━┓ ┏━━━┓   ┏━━━┓ ┏━━━┓ ┏━┓┏━┓ ┏━━━┓ ┏┓    ┏━━━┓ ┏━━━━┓ ┏━━━┓ ┏━━━┓   ┏━━━┓ ┏━━━┓ ┏━━━┓ ┏┓ ┏┓ ┏━━━┓ ┏━━━┓ ┏━━━━┓ ┏━━━┓
# ┃┃ ┃┃ ┃┏━┓┃ ┗┓┏┓┃ ┃┏━┓┃ ┃┏┓┏┓┃ ┃┏━━┛   ┃┏━┓┃ ┃┏━┓┃ ┃ ┗┛ ┃ ┃┏━┓┃ ┃┃    ┃┏━━┛ ┃┏┓┏┓┃ ┃┏━━┛ ┗┓┏┓┃   ┃┏━┓┃ ┃┏━━┛ ┃┏━┓┃ ┃┃ ┃┃ ┃┏━━┛ ┃┏━┓┃ ┃┏┓┏┓┃ ┃┏━┓┃
# ┃┃ ┃┃ ┃┗━┛┃  ┃┃┃┃ ┃┃ ┃┃ ┗┛┃┃┗┛ ┃┗━━┓   ┃┃ ┗┛ ┃┃ ┃┃ ┃┏┓┏┓┃ ┃┗━┛┃ ┃┃    ┃┗━━┓ ┗┛┃┃┗┛ ┃┗━━┓  ┃┃┃┃   ┃┗━┛┃ ┃┗━━┓ ┃┃ ┃┃ ┃┃ ┃┃ ┃┗━━┓ ┃┗━━┓ ┗┛┃┃┗┛ ┃┗━━┓
# ┃┃ ┃┃ ┃┏━━┛  ┃┃┃┃ ┃┗━┛┃   ┃┃   ┃┏━━┛   ┃┃ ┏┓ ┃┃ ┃┃ ┃┃┃┃┃┃ ┃┏━━┛ ┃┃ ┏┓ ┃┏━━┛   ┃┃   ┃┏━━┛  ┃┃┃┃   ┃┏┓┏┛ ┃┏━━┛ ┃┗━┛┃ ┃┃ ┃┃ ┃┏━━┛ ┗━━┓┃   ┃┃   ┗━━┓┃
# ┃┗━┛┃ ┃┃    ┏┛┗┛┃ ┃┏━┓┃   ┃┃   ┃┗━━┓   ┃┗━┛┃ ┃┗━┛┃ ┃┃┃┃┃┃ ┃┃    ┃┗━┛┃ ┃┗━━┓   ┃┃   ┃┗━━┓ ┏┛┗┛┃   ┃┃┃┗┓ ┃┗━━┓ ┗━━┓┃ ┃┗━┛┃ ┃┗━━┓ ┃┗━┛┃   ┃┃   ┃┗━┛┃
# ┗━━━┛ ┗┛    ┗━━━┛ ┗┛ ┗┛   ┗┛   ┗━━━┛   ┗━━━┛ ┗━━━┛ ┗┛┗┛┗┛ ┗┛    ┗━━━┛ ┗━━━┛   ┗┛   ┗━━━┛ ┗━━━┛   ┗┛┗━┛ ┗━━━┛    ┗┛ ┗━━━┛ ┗━━━┛ ┗━━━┛   ┗┛   ┗━━━┛
#
#In a nutshell: this process goes through each email in the Master Data Requests folder by assigned catagory (person) and converts the email message into text.
#Searches the text to extract the info needed: request number, completed URL, original file name.
#Open the original excel file, updates the status to complete, pastes the completed URL, saves and closes.
#Emails person about each completion or error so that person can update their tasks or inspect the errors manually.
#
#Jan 2018
#jemurra
#############################################################################################################

import win32com.client as win32
import requests
from requests_ntlm import HttpNtlmAuth
import re
from datetime import datetime
from logins import *

#setting up Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
olaccount = outlook.Folders.Item("MDSEMDS@email.com")
inbox = olaccount.Folders("Inbox")
reqs = inbox.Folders("Folder Name Here")

#Sharepoint Beginning Path
MDSSP = "http://URL_HERE"
dt = datetime.now()

#Excel app
excel = win32.Dispatch("Excel.Application")
excel.Visible = 0

#messages and tasks in Outlook
msgs = reqs.Items

#Array
mdbox = []

#Function to get the completed url link and text for completed file in MD email
def getURL(b):
	return re.search("(?P<url>https?://[^\s]+)", b).group("url")

#Function to get Title of original file excel file, used to append onto MDSE sharepoint URL to open the excel file. Check email first, othewise get excel file title from MDS waypoint log
def getORGp(of, orn):
	for word in of.split("	"):
		if "xlsm" in word:
			if "http://" not in word:
				orgf = word.lstrip()
				return orgf
		else:
			urlB = 'http://teamsites_here'
			exf = urlB + orn[-10:]
			rSP = requests.get(exf, auth=HttpNtlmAuth(usrn, pswd))
			for word in rSP.text.split("	"):
				for subw in word.split(","):
					for nn in subw.split('"'):
						if "xlsm" in nn:
							return(nn)

def updateExcel(b, sbj):
	print("Updating " + sbj)
	exSP = str(MDSSP) + str(getORGp(b,sbj))

	try:
		myex = excel.Workbooks.Open(exSP)
		for p in myex.ContentTypeProperties:
			if p.Name == "Status":
				p.Value = "Complete"
			if p.Name == "Master Data Completed File":
				p.Value = getURL(b)
		myex.Save()
		myex.Close()
		print("Done updating " + sbj[-10:])

		mdbox.append(["succ", sbj])

	except Exception as e:
		print(e)
		mdbox.append(["err",  sbj])

#Function to get empty arrays if data is in them and get emails in box based on assigned category
def getEmails(per):
  mdbox.clear()
  print("Updating " +str(per) + " requests")
  for a in msgs:
  	if str(a.Categories) == per:
  		body = a.Body
  		sbjt = a.Subject
  		reqn = sbjt[-10:]
  		if 'has been opened' in sbjt:
  			mdbox.append(["sub", sbjt])
  		elif 'D97 ' in sbjt:
  			mdbox.append(["grg", sbjt])
  		elif 'D00' in sbjt:
  			mdbox.append(["err", "Potential MP Rollup, update manually: " + sbjt])
  		elif 'http:' not in body:
  			mdbox.append(["err", "no completed file: " + str(sbjt)])
  		else:
  			#mdbox.append(["succ", sbjt])
  			updateExcel(body, sbjt)

def delmsgs(msub):
	for msg in msgs:
		if msub == msg.Subject:
			print("Deleting " + str(msg.Subject))
			msg.Delete()

def getComments(person):
	for msg in msgs:
		if msg.Categories == person:
			for word in msg.Body.split("	"):
				if 'Attention' in word:
					zebra = word.partition("Attention")
					mdbox.append(["cmt", msg.Subject[-10:] + " -" + zebra[2]])

#Function to send email to each person to let them know what was updated or not updated
def sendEmail(emTo):
	AMs = []
	ACs = []
	Othrs = []
	Opns = []
	Ers = []
	Cmts = []

	for tpe in mdbox:
		if "opened" in tpe[1]:
			Opns.append(tpe[1])
		elif "err" in tpe[0]:
			Ers.append(tpe[1])
		elif "gr" in tpe[0]:
			Othrs.append(tpe[1])
		elif tpe[0] == "cmt":
			Cmts.append(tpe[1])
		elif "Article Maintain" in tpe[1]:
			AMs.append(tpe[1])
		elif "Article Create" in tpe[1]:
			ACs.append(tpe[1])
		else:
			Othrs.append(tpe[1])

	outlook = win32.Dispatch('outlook.application')
	mail = outlook.CreateItem(0)
	mail.To = emTo
	mail.Subject = 'Updated Requests %s-%s-%s %s'  %(dt.month, dt.day, dt.year, dt.strftime("%H:%M"))
	mail.body = 'The following Master Data Requests have been marked as completed and updated with the completed MD file (except Submissions):' + '\n' + '\n' + \
	"Total Emails in box - " + str((len(mdbox))-len(Cmts)) + '\n' + '\n' + \
	("Article Creates - " + str(len(ACs))  + '\n' + str('\n'.join(ACs)) + '\n' + '\n' if len(ACs) > 0 else "") + \
	("Article Maintains - " + str(len(AMs)) + '\n' + str('\n'.join(AMs)) + '\n' + '\n' if len(AMs) > 0 else "") + \
	("Others - " + str(len(Othrs)) + '\n' + str('\n'.join(Othrs))  + '\n' + '\n' if len(Othrs) > 0 else "") + \
	("Errors/Anomalies - " + str(len(Ers)) + '\n' + str('\n'.join(Ers))  + '\n' + '\n' if len(Ers) > 0 else "") + \
	("Submissions - " + str(len(Opns))  + '\n' + str('\n'.join(Opns))  + '\n' + '\n' if len(Opns) > 0 else "") + \
	("Comments" + '\n' + str('\n'.join(Cmts)) if len(Cmts) > 0 else "")
	mail.send

#List of people to run the process on and their email
reqCats = ['Amy', 'Bennet', 'Brooke'] #names here
mdsemails = ['name1@nope.com', 'name2@nope.com', 'name3@nope.com'] #emails here

justOne = True
runAll = False

if justOne:
	per = 'Me'
	peridx = reqCats.index(per)
	getEmails(per)
	getComments(per)
	print("Deleting Completed Emails")
	for a in mdbox:
		if a[0] != 'err':
			delmsgs(a[1])
	sendEmail(mdsemails[peridx])

if runAll:
	for per in reqCats:
		perIDX = reqCats.index(per)
		getEmails(per)
		getComments(per)
		print("Deleting Completed Emails")
		for a in mdbox:
			if a[0] != 'err':
				delmsgs(a[1])
		sendEmail(mdsemails[perIDX])

excel.Quit()
print("######### Finished #########")
