import tbapy
import xlwt
from xlwt import Workbook


tba = tbapy.TBA('3WEGx9vYItqOFTwzOjip8LwmwQ4VpCJCfS0jlPdlqOP76XFkcEh3x66i2HzcRrq5')


def eventAverage(event):
    dictionary = tba.event_insights(event)
    return round(dictionary['qual']['average_score'])


def winScore(event):
    dictionary = tba.event_insights(event)
    return round(dictionary['qual']['average_win_score'])


def teamAverage(team, event):
    myList = list()
    dictionary = tba.team_matches(team, event)
    mySum = 0
    count = 0
    for i in dictionary:
        if ('frc' + str(team)) in (i['alliances']['blue']['team_keys']):
            mySum += (i['alliances']['blue']['score'])
            count += 1
        else:
            mySum += (i['alliances']['red']['score'])
            count += 1
    return round(mySum / count)


def rocket(team, event, sheet,column):
    lowerRocket = list()
    middleRocket = list()
    upperRocket = list()
    dictionary = tba.team_matches(team, event)
    scoresList=[]

    for i in dictionary:
        breakDownBlue = i['score_breakdown']['blue']
        breakDownRed = i['score_breakdown']['red']
        if ('frc' + str(team)) in (i['alliances']['red']['team_keys']):
            lowerRocket.append(breakDownRed['lowLeftRocketNear'])
            lowerRocket.append(breakDownRed['lowLeftRocketFar'])
            lowerRocket.append(breakDownRed['lowRightRocketNear'])
            lowerRocket.append(breakDownRed['lowRightRocketFar'])

            middleRocket.append(breakDownRed['midLeftRocketNear'])
            middleRocket.append(breakDownRed['midLeftRocketFar'])
            middleRocket.append(breakDownRed['midRightRocketNear'])
            middleRocket.append(breakDownRed['midRightRocketFar'])

            upperRocket.append(breakDownRed['topLeftRocketNear'])
            upperRocket.append(breakDownRed['topLeftRocketFar'])
            upperRocket.append(breakDownRed['topRightRocketNear'])
            upperRocket.append(breakDownRed['topRightRocketFar'])
        else:
            lowerRocket.append(breakDownBlue['lowLeftRocketNear'])
            lowerRocket.append(breakDownBlue['lowLeftRocketFar'])
            lowerRocket.append(breakDownBlue['lowRightRocketNear'])
            lowerRocket.append(breakDownBlue['lowRightRocketFar'])

            middleRocket.append(breakDownRed['midLeftRocketNear'])
            middleRocket.append(breakDownRed['midLeftRocketFar'])
            middleRocket.append(breakDownRed['midRightRocketNear'])
            middleRocket.append(breakDownRed['midRightRocketFar'])

            upperRocket.append(breakDownRed['topLeftRocketNear'])
            upperRocket.append(breakDownRed['topLeftRocketFar'])
            upperRocket.append(breakDownRed['topRightRocketNear'])
            upperRocket.append(breakDownRed['topRightRocketFar'])

    lowerRocketScore=((round((lowerRocket.count('Panel') / len(lowerRocket)) * 100))+(round((lowerRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))/2
    middleRocketScore=((round((middleRocket.count('Panel') / len(lowerRocket)) * 100))+(round((middleRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))/2
    upperRocketScore=((round((upperRocket.count('Panel') / len(lowerRocket)) * 100))+(round((upperRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))/2
    rocketOVR=lowerRocketScore+(2*middleRocketScore)+(3 * upperRocketScore)

    sheet.write(4, column, (round((lowerRocket.count('Panel') / len(lowerRocket)) * 100)))
    sheet.write(5, column, (round((lowerRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))
    sheet.write(6, column, (round((middleRocket.count('Panel') / len(lowerRocket)) * 100)))
    sheet.write(7, column, (round((middleRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))
    sheet.write(8, column, (round((upperRocket.count('Panel') / len(lowerRocket)) * 100)))
    sheet.write(9, column, (round((upperRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))
    sheet.write(10,column, (round(rocketOVR)))

def teamReport(team, event,sheet,column):
    average = teamAverage(team, event)
    sheet.write(0, column, (str(team)))
    sheet.write(1, column, (average))
    sheet.write(2, column, average-(eventAverage(event)))
    rocket(team, event, sheet, column)


def eventReport(event):
    wb=Workbook()
    sheet1=wb.add_sheet('sheet1')
    sheet1.write(0, 0, "Team Number: ")
    sheet1.write(1, 0, "Team Average Score: ")
    sheet1.write(2, 0, "Points above/below Event Average:")
    sheet1.write(4, 0, "Lower Rocket Panel Percentage: ")
    sheet1.write(5, 0, "Lower Rocket Panel and Cargo Percentage: ")
    sheet1.write(6, 0, "Middle Rocket Panel Percentage: ")
    sheet1.write(7, 0, "Middle Rocket Panel and Cargo Percentage: ")
    sheet1.write(8, 0, "Upper Rocket Panel Percentage: ")
    sheet1.write(9, 0, "Upper Rocket Panel and Cargo Percentage: " )
    sheet1.write(10,0, "OVR Rocket Rating:")
    for i in range(1,len(event_teams(event))+1):
        teamReport(event_teams(event)[i-1],event,sheet1,i)

    wb.save(str(event)+".xls")

def event_teams(event):
    list = tba.event_teams(event)
    myList=[]

    for i in list:
        teamKey=i.key
        team=str(teamKey).replace("frc","")
        number=int(team)
        myList.append(number)
    myList.sort()
    return myList

