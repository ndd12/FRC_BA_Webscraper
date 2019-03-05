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


def rocket(team, event, sheet):
    lowerRocket = list()
    middleRocket = list()
    upperRocket = list()
    dictionary = tba.team_matches(team, event)

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

    sheet.write(2, 0, "Lower Rocket Panel Percentage: ")
    sheet.write(2, 1, str(round((lowerRocket.count('Panel') / len(lowerRocket)) * 100)) + "%")

    sheet.write(3, 0, "Lower Rocket Panel and Cargo Percentage: ")
    sheet.write(3, 1, str(round((lowerRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)) + "%")

    sheet.write(4, 0, "Middle Rocket Panel Percentage: ")
    sheet.write(4, 1, str(round((middleRocket.count('Panel') / len(lowerRocket)) * 100)) + "%")

    sheet.write(5, 0, "Middle Rocket Panel and Cargo Percentage: ")
    sheet.write(5, 1, str(round((middleRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)) + "%")

    sheet.write(6, 0, "Upper Rocket Panel Percentage: ")
    sheet.write(6, 1,  str(round((upperRocket.count('Panel') / len(lowerRocket)) * 100)) + "%")

    sheet.write(7, 0, "Upper Rocket Panel and Cargo Percentage: " )
    sheet.write(7, 1, str(round((upperRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)) + "%")

def teamReport(team, event):
    wb=Workbook()
    sheet1=wb.add_sheet('sheet1')
    average = teamAverage(team, event)
    sheet1.write(0, 0, "Team Number: ")
    sheet1.write(0, 1, (str(team)))
    sheet1.write(1,0,"Team Average Score: ")
    sheet1.write(1,1,(str(average) + "\n"))

    rocket(team, event, sheet1)

    wb.save("/Users/noahdouglas/Desktop/Team #"+str(team)+".xls")

def teamMatches(team, event):
    dictionary = tba.team_matches(team, event)
    for i in dictionary:
        print(i['score_breakdown']['blue'].keys())
