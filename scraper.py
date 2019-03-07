# check if needed modules are installed
try:
    import tbapy
except ImportError:
    print('Error, Module tbapy is required')

try:
    import xlwt
except ImportError:
    print('Error, Module xlwt is required')

# pass in key generated within blue alliance account
tba = tbapy.TBA('3WEGx9vYItqOFTwzOjip8LwmwQ4VpCJCfS0jlPdlqOP76XFkcEh3x66i2HzcRrq5')

# global styles used to write to spreadsheet
styleBold = xlwt.XFStyle()
bold = xlwt.Font()
bold.bold = True
bold.height = 240
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
styleBold.alignment = alignment
styleBold.font = bold

style = xlwt.XFStyle()
font = xlwt.Font()
font.height = 240
style.alignment = alignment
style.font = font


# Find the average score for an event
def event_average(event):
    dictionary = tba.event_insights(event)
    return round(dictionary['qual']['average_score'])


# calculate the average points a team scored at a given event
def team_average(team, event):
    # import information about the matches a team had at an event
    dictionary = tba.team_matches(team, event)
    # variables for calculating averages
    mySum = 0
    count = 0
    # for loop to iterate through a team's matches
    for i in dictionary:
        # determine what alliance team was on in the match and use appropriate score in calculation
        if ('frc' + str(team)) in (i['alliances']['blue']['team_keys']):
            mySum += (i['alliances']['blue']['score'])
            count += 1
        else:
            mySum += (i['alliances']['red']['score'])
            count += 1
    # calculate and return the average has a whole number
    return round(mySum / count)


# note: rocket() is called within team_report(), and not directly called within event_report()
def rocket(team, event, sheet,column):
    styleBold = xlwt.XFStyle()
    bold = xlwt.Font()
    bold.bold = True
    bold.height = 240
    styleBold.font = bold

    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.height = 240
    style.font = font

    # 3 lists to hold the results of every match related to that specific part of the rocket
    lowerRocket = list()
    middleRocket = list()
    upperRocket = list()
    # dictionary used to hold information about teams matches
    dictionary = tba.team_matches(team, event)
    # iterate through every match for given team
    for i in dictionary:
        # separate blue and red information for current match
        breakDownBlue = i['score_breakdown']['blue']
        breakDownRed = i['score_breakdown']['red']

        # if team is on red for current match, read information on red alliance, otherwise, read in blue information
        if ('frc' + str(team)) in (i['alliances']['red']['team_keys']):
            # Read in information for all low rocket zones in the match, then do the same for middle slot and high slot
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
    # algorithm for estimating a teams ability at each rocket level, and then an overall rocket rating
    lowerRocketScore=((round((lowerRocket.count('Panel') / len(lowerRocket)) * 100)) + (round((lowerRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))/2
    middleRocketScore=((round((middleRocket.count('Panel') / len(lowerRocket)) * 100))+(round((middleRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))/2
    upperRocketScore=((round((upperRocket.count('Panel') / len(lowerRocket)) * 100))+(round((upperRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)))/2
    rocketOVR=lowerRocketScore+(2*middleRocketScore)+(3 * upperRocketScore)

    # once information has been calculated, write into the sheet for event_report
    sheet.write(4, column, (round((lowerRocket.count('Panel') / len(lowerRocket)) * 100)), style)
    sheet.write(5, column, (round((lowerRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)), style)
    sheet.write(6, column, (round((middleRocket.count('Panel') / len(lowerRocket)) * 100)), style)
    sheet.write(7, column, (round((middleRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)), style)
    sheet.write(8, column, (round((upperRocket.count('Panel') / len(lowerRocket)) * 100)), style)
    sheet.write(9, column, (round((upperRocket.count('PanelAndCargo') / len(lowerRocket)) * 100)), style)
    sheet.write(10,column, (round(rocketOVR)), style)


# return a sorted list of the team numbers competing at a given event
def event_teams(event):
    list = tba.event_teams(event)
    myList=[]
    for i in list:
        teamKey = i.key
        team = str(teamKey).replace("frc", "")
        number = int(team)
        myList.append(number)
    myList.sort()
    return myList


# write the information for given team into the event report
def team_report(team, event, sheet, column):
    average = team_average(team, event)
    sheet.write(0, column, (str(team)), styleBold)
    sheet.write(1, column, average, style)
    sheet.write(2, column, average-(event_average(event)), style)
    rocket(team, event, sheet, column)
    end_game_location(team,event,sheet,column)

def event_report(event):

    # create new spreadsheet in project folder
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet('sheet1')
    first_col=sheet1.col(0)
    first_col.width = 256*45
    sheet1.set_panes_frozen(True)
    sheet1.set_horz_split_pos(1)
    sheet1.set_vert_split_pos(1)
    # hard-code row names
    sheet1.write(0, 0, "Team Number: ", styleBold)
    sheet1.write(1, 0, "Team Average Score: ", styleBold)
    sheet1.write(2, 0, "Points above/below Event Average:", styleBold)
    sheet1.write(4, 0, "Lower Rocket Panel Percentage: ", styleBold)
    sheet1.write(5, 0, "Lower Rocket Panel and Cargo Percentage: ", styleBold)
    sheet1.write(6, 0, "Middle Rocket Panel Percentage: ", styleBold)
    sheet1.write(7, 0, "Middle Rocket Panel and Cargo Percentage: ", styleBold)
    sheet1.write(8, 0, "Upper Rocket Panel Percentage: ", styleBold)
    sheet1.write(9, 0, "Upper Rocket Panel and Cargo Percentage: ", styleBold)
    sheet1.write(10, 0, "OVR Rocket Rating:", styleBold)
    sheet1.write(12, 0, "HAB Level 1 Finish Percentage:", styleBold)
    sheet1.write(13, 0, "HAB Level 2 Finish Percentage:", styleBold)
    sheet1.write(14, 0, "HAB Level 3 Finish Percentage:", styleBold)

    # generate a new team report for every team at the given event
    for i in range(1, len(event_teams(event))+1):
        team_report(event_teams(event)[i-1], event, sheet1, i)

    # save the event report after all team reports have been calculated
    wb.save(str(event)+".xls")


# given a team and an event, find how robot finished in respect to HAB zones for every match
def end_game_location(team, event, sheet, column):
    # a list to hold all of the final locations
    location = []
    # iterate through every match the team has at the given event
    for i in tba.team_matches(team, event):
        # If the team is on the red alliance, look at the score breakdown for the red alliane or blue otherwise
        if ('frc' + str(team)) in (i['alliances']['red']['team_keys']):
            teams=i['alliances']['red']['team_keys']
            # determine if robot is robot 1 , 2, or 3 in the given match and record the corresponding value from the api
            if ("frc"+str(team)) == teams[0]:
                location.append(i['score_breakdown']['red']['endgameRobot1'])
            elif ("frc"+str(team)) == teams[1]:
                location.append(i['score_breakdown']['red']['endgameRobot2'])
            elif ("frc"+str(team)) == teams[2]:
                location.append(i['score_breakdown']['red']['endgameRobot3'])

        elif ('frc' + str(team)) in (i['alliances']['blue']['team_keys']):
            teams = i['alliances']['blue']['team_keys']
            if ("frc"+str(team)) == teams[0]:
                location.append(i['score_breakdown']['blue']['endgameRobot1'])
            elif ("frc"+str(team)) == teams[1]:
                location.append(i['score_breakdown']['blue']['endgameRobot2'])
            elif ("frc"+str(team)) == teams[2]:
                location.append(i['score_breakdown']['blue']['endgameRobot3'])
    # Get values for what percent of the matches teams robot is in the different positions of the HAB Zone
    levelOnePercentage = (location.count("HabLevel1")/len(location)) * 100
    levelTwoPercentage = (location.count("HabLevel2")/len(location)) * 100
    levelThreePercentage = (location.count("HabLevel3")/len(location)) * 100

    # write values that were just calculated to the sheet used in the function call
    sheet.write(12, column, round(levelOnePercentage), style)
    sheet.write(13, column, round(levelTwoPercentage), style)
    sheet.write(14, column, round(levelThreePercentage), style)


def insights(team, event):
    for i in tba.team_matches(8, "2019cadm"):
        if ('frc' + str(team)) in (i['alliances']['red']['team_keys']):
            print (i['alliances']['red']['team_keys'])
            print (i['score_breakdown']['red'])


