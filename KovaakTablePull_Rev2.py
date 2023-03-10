import json
import math
import requests
import xlwt
from xlwt import Workbook

# INPUT IN SCENARIO NAMES
SCENARIO_NAMES = ['Pasu Voltaic Easy', 'B180 Voltaic Easy']
Excel_File_Name = 'Test'

# ARRAY SETUP
Score_Array = [[] for ii in range(0, len(SCENARIO_NAMES))]
Name_Array = [[] for iii in range(0, len(SCENARIO_NAMES))]
Account_Name = []
Account_Name_Unique = []
Leaderboard_ID = [0] * len(SCENARIO_NAMES)

# REQUEST SCENARIO PATH ONE TIME TO GET AMOUNT OF PAGES ON THE SCENARIOS PAGE
r = requests.get('https://kovaaks.com/webapp-backend/scenario/popular?page=0&max=100')
r = json.loads(r.text)
Max_Page = math.floor(r["total"]/100)

# ITERATE THOUGH ALL PAGES ON THE SCENARIOS PAGE TO FIND MATCH WITH SCENARIOS (EACH PAGE HAS A MAX of 100 ROWS)
Count = 0
for i in range(0, Max_Page + 1):
    r = requests.get('https://kovaaks.com/webapp-backend/scenario/popular?page=' + str(i) + '&max=100')
    r = json.loads(r.text)
    ra = r["data"]

    # LOOK FOR MATCH WITH THE SCENARIO NAMES ON PAGE
    for ii in range(0, len(ra)):
        test = (ra[ii])

        # IF MATCH SEND SCENARIO ID VALUE TO LEADERBOARD ID ARRAY
        if test["scenarioName"] in SCENARIO_NAMES:
            Find_Location = SCENARIO_NAMES.index(test["scenarioName"])
            Leaderboard_ID[Find_Location] = test["leaderboardId"]
            print("Scenario ID Found for: " + SCENARIO_NAMES[Find_Location] + ". " + str(Count+1) + " of " + str(len(SCENARIO_NAMES)) + ".")
            Count = Count + 1

    # EXIT LOOP IF ALL LEADERBOARD IDs HAVE BEEN FOUND
    if Count >= len(SCENARIO_NAMES):
        break

# ITERATE THROUGH EACH LEADERBOARDS
for i in range(0, len(SCENARIO_NAMES)):

    # FILE REQUEST PATH ONE TIME TO GET AMOUNT OF PAGES ON LEADERBOARD
    r = requests.get('https://kovaaks.com/webapp-backend/leaderboard/scores/global?leaderboardId=' + str(Leaderboard_ID[i]) + '&page=0&max=100')
    r = json.loads(r.text)
    Max_Page = math.floor(r["total"]/100)

    # ITERATE THOUGH ALL PAGES ON THE API LEADERBOARD (EACH PAGE HAS A MAX of 100 ROWS)
    for ii in range(0, Max_Page + 1):
        r = requests.get('https://kovaaks.com/webapp-backend/leaderboard/scores/global?leaderboardId=' + str(Leaderboard_ID[i]) + '&page=' + str(ii) + '&max=100')
        r = json.loads(r.text)
        ra = r["data"]
        print("Leaderboard: " + str(i+1) + " of " + str(len(SCENARIO_NAMES)) + ".  Page: " + str(ii) + " of " + str(Max_Page) + ".")

        # SEND RELEVANT DATA FROM EACH PAGE TO THREE ARRAYS (SCORE ARRAY, NAME ARRAY, MAIN NAME ARRAY)
        for iii in range(0, len(ra)):
            test = (ra[iii])
            Score_Array[i].append(test["score"])
            Name_Array[i].append(test["steamAccountName"])
            Account_Name.append(test["steamAccountName"])

# REMOVE DUPLICATE USERNAMES IN MAIN NAME ARRAY
for i in Account_Name:
    if i not in Account_Name_Unique:
        Account_Name_Unique.append(i)

# CREATE EXCEL FILE AND CREATE COLUMN HEADERS
wb = Workbook()
sheet1 = wb.add_sheet('Combined Stats')
sheet1.write(0, 0, 'Name', xlwt.easyxf('font: bold 1'))
for i in range(0, len(SCENARIO_NAMES)):
    sheet1.write(0, i + 1, 'Scenario: ' + SCENARIO_NAMES[i], xlwt.easyxf('font: bold 1'))
    
# ITERATE THROUGH EVERY UNIQUE USERNAME TO SEND TO EXCEL
for i in range(0, len(Account_Name_Unique)):
    print("Send to Excel: " + str(i + 1) + " of " + str(len(Account_Name_Unique)) + ".")
    try:  # Sometimes excel write errors here, so this
        sheet1.write(i + 1, 0, Account_Name_Unique[i])
    except:
        pass

    # FOR EVERY USERNAME ITERATE THROUGH ALL LEADERBOARDS TO FIND A SCORE
    for ii in range(0, len(SCENARIO_NAMES)):

        # IF MATCH IS FOUND SEND SCORE TO EXCEL SHEET
        if Account_Name_Unique[i] in Name_Array[ii]:
            Find_Location = Name_Array[ii].index(Account_Name_Unique[i])
            try:  # Sometimes excel write errors here, so this
                sheet1.write(i + 1, ii + 1, Score_Array[ii][Find_Location])
            except:
                pass

# SAVE EXCEL SHEET
wb.save('Leaderboard_Pull_For_' + Excel_File_Name + '.xls')
