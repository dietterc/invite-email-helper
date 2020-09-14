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

startingPoint = int(input("What line are we starting at? ")) - 1

print(names[startingPoint])