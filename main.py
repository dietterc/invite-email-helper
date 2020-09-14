import sys
import os
import pyHook, pythoncom

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

SUBJECT = "UManitoba Computer Science Discord Invitation"


def main():

    startIndex = int(input("What line are we starting at? ")) - 1

    emailBodies = generateEmails(startIndex)

    newInvites = getDiscordInvites(len(emailBodies))

    finalEmails = compileEmails(emailBodies,newInvites)

    for i in finalEmails:
        i.printEmail()

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
            
    
    template = open("template.txt","r").read()
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