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


def main():

    startIndex = int(input("What line are we starting at? ")) - 1

    emailBodies = generateEmails(startIndex)

    newInvites = getDiscordInvites(len(emailBodies))

    

def generateEmails(startIndex):
    global names
    global emails
    global flaggedIndexes
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
    for i in range(len(noDuplicates)):
        currentIndex = noDuplicates[i]

        if emails[currentIndex].endswith("@myumanitoba.ca"):
            goodEntries.append(currentIndex)
        else:
            flaggedIndexes.append(currentIndex)
            
    
    template = open("template.txt","r").read()
    emailTemplates = []

    for i in range(len(goodEntries)):

        firstName = names[goodEntries[i]].split(" ")[0]

        currentTemplate = template.format(name = firstName, invite = "{invite}")
        emailTemplates.append(currentTemplate)

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
                for i in range(len(invites)):
                    if(invites[i] == invite):
                        found = True
                
                #check new invites too
                for i in range(len(newInvites)):
                    if(newInvites[i] == invite):
                        found = True
                    
                if(found):
                    print("Invalid invite - Already in use")
                else:
                    newInvites.append(invite)
                    count += 1

            else:
                print("Invalid invite - Invite must start with https://discord.gg/")

        print("\nInvites received:")
        for i in range(len(newInvites)):
            print(newInvites[i])
        print("")

        while(confirm != 'y' and confirm != 'n'):
            confirm = input("Are the above invites okay? (y/n) ")

        if confirm == 'y':
            return newInvites
        else:
            confirm = ""


class Email:
    def __init__(self, dest, subject, body):
        self.dest = dest
        self.subject = subject
        self.body = body

main()