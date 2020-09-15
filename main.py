import sys
import os
import pyHook, pythoncom
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#Google api stuff
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#open the responses spread-sheet
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

responsesSheet = client.open("Responses to discord signup").sheet1

#Get info into lists
names = responsesSheet.col_values(2)
emails = responsesSheet.col_values(3)
usernames = responsesSheet.col_values(5)
invites = responsesSheet.col_values(6)


#list of flagged indexes (form entries that threw some kind of error)
flaggedIndexes = []
goodIndexes = []
finalEmails = []

SUBJECT = "UManitoba Computer Science Discord Invitation"

viewIndex = -1

def main():
    global finalEmails

    startIndex = int(input("What line are we starting at? ")) - 1

    emailBodies = generateEmails(startIndex)

    newInvites = getDiscordInvites(len(emailBodies))

    finalEmails = compileEmails(emailBodies,newInvites)

    #clear the terminal
    os.system('cls')

    print("Use the left and right arrow keys to scroll between email previews.")

    hm = pyHook.HookManager()
    hm.KeyDown = OnKeyboardEvent
    hm.HookKeyboard()

    while True:
        pythoncom.PumpWaitingMessages()
        if(len(finalEmails) == 0):
            break

    flagIndexes()
    print("Finished.")

def OnKeyboardEvent(event): 
    global viewIndex
    global finalEmails
    global goodIndexes
    key = event.Key

    if key == 'Left':
        if(viewIndex > 0):
            viewIndex -= 1
            os.system('cls')
            finalEmails[viewIndex].printEmail()
            printFooter(viewIndex)
            
    elif key == 'Right':
        if(viewIndex != len(finalEmails) -1):
            viewIndex += 1
            os.system('cls')
            finalEmails[viewIndex].printEmail()
            printFooter(viewIndex)
    elif key == 'Q':
        flagIndexes()
        exit(0)
    elif key == "Return":
        sendEmail(finalEmails[viewIndex],goodIndexes[viewIndex])
    elif key == "Delete":
        flaggedIndexes.append(goodIndexes[viewIndex])
        finalEmails.remove(finalEmails[viewIndex])
        goodIndexes.remove(goodIndexes[viewIndex])
        os.system('cls')

        print("\nIndex flagged.\n-----------------\n")
        print("Use the arrow keys to navigate")
        viewIndex = -1

    return True

def sendEmail(email,index):
    global finalEmails
    global goodIndexes
    global viewIndex

    os.system('cls')

    #get the discord invite from the email
    bodyWords = email.body.split("\n")
    invite = ""
    for i in bodyWords:
        if i.startswith("https://discord.gg/"):
            invite = i

    #prep and send the email
    sender_email = 'cssadiscordinvites@gmail.com'
    gmail_password = open("gmail_password.txt","r").read()

    #just me for now
    receiver_email = "dietterc@myumanitoba.ca"

    message = MIMEMultipart("alternative")
    message["Subject"] = email.subject
    message["From"] = "UofM CS Discord Form <" + sender_email + ">"
    message["To"] = receiver_email

    text = open("template_plain.txt","r").read().format(d_invite = invite)

    html = email.body

    message.attach(MIMEText(text, "plain"))
    message.attach(MIMEText(html, "html"))

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(sender_email, gmail_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.close()

        print("\nEmail sent!\n-----------------\n")
        print("Use the arrow keys to navigate")

        #if it sent then add the discord invite to the spreadsheet
        responsesSheet.update_cell(index+1,6,invite)

        viewIndex = -1
        finalEmails.remove(email)
        goodIndexes.remove(index)

    except:
        print("\nSomething went wrong... Email not sent.\n-----------------\n")
        print("Use the arrow keys to navigate")
        viewIndex = -1


def flagIndexes():
    global flaggedIndexes

    for i in flaggedIndexes:
        responsesSheet.update_cell(i+1,7,"FLAGGED")



def printFooter(index):
    print("Email {i} of {totalIndex}\n".format(i=index + 1,totalIndex=len(finalEmails)))
    print("Use the arrow keys to navigate")
    print("ENTER to confirm and send email")
    print("DELETE to flag (and not send)")


def generateEmails(startIndex):
    global names
    global emails
    global flaggedIndexes
    global goodIndexes

    emailCount = len(names) - startIndex
    noDuplicates = []
    goodEntries = []

    #for each user starting at the given index, check if they have already signed up for the form. 
    for i in range(startIndex, startIndex + emailCount):
        found = False

        for j in range(startIndex + (i - startIndex)):
            if(emails[i] == emails[j]):
                found = True

        if not found:
            noDuplicates.append(i)
        else:
            #duplicate entry found
            flaggedIndexes.append(i)

    #Check if its a myumanitoba email 
    for i in noDuplicates:

        if emails[i].endswith("@myumanitoba.ca"):
            goodEntries.append(i)
        else:
            flaggedIndexes.append(i)
            
    
    template = open("template_html.txt","r").read()
    emailTemplates = []

    for i in goodEntries:

        firstName = names[i].split(" ")[0]

        currentTemplate = template.format(name = firstName, invite = "{invite}")
        emailTemplates.append(currentTemplate)

    goodIndexes = goodEntries
    return emailTemplates


def getDiscordInvites(numInvites):

    global invites
    confirm = ""

    while(confirm != 'y'):
        newInvites = []

        count = 0
        while(count < numInvites):
            invite = input("Enter discord invite {curr}/{total}:\n".format(curr = count, total = numInvites))
            if(invite.startswith("https://discord.gg/")):
                found = False
                for i in invites:
                    if(i == invite):
                        found = True
                
                #check new invites too
                for i in newInvites:
                    if(i == invite):
                        found = True
                    
                if(found):
                    print("Invalid invite - Already in use")
                else:
                    newInvites.append(invite)
                    count += 1

            else:
                print("Invalid invite - Invite must start with https://discord.gg/")

        print("\nInvites received:")
        for i in newInvites:
            print(i)
        print("")

        while(confirm != 'y' and confirm != 'n'):
            confirm = input("Are the above invites okay? (y/n) ")

        if confirm == 'y':
            return newInvites
        else:
            confirm = ""


def compileEmails(bodies,newInvites):
    global emails
    global goodIndexes

    finalEmails = []
    for i in range(len(goodIndexes)):
        body = bodies[i].format(invite = newInvites[i])
        finalEmails.append(Email(emails[goodIndexes[i]],SUBJECT,body))

    return finalEmails



class Email:
    def __init__(self, dest, subject, body):
        self.dest = dest
        self.subject = subject
        self.body = body

    def printEmail(self):
        print("TO:\n" + self.dest + "\n")
        print("SUBJECT:\n" + self.subject + "\n")
        print("BODY:\n" + self.body + "\n-----------------")


main()