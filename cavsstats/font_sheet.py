import xlwt
import xlrd

workbook = xlwt.Workbook()
fontsheet = workbook.add_sheet('Font Sheet')
leaders = workbook.add_sheet('Leaders Sheet')
fontsheet.write_merge(0, 0, 0, 11, 'PREGAME')
fontsheet.write(1, 0, '#')
fontsheet.write(1, 1, 'AWAY PLAYER')
fontsheet.write(1, 2, 'Career')
fontsheet.write(1, 3, 'This Season')
fontsheet.write(1, 4, 'Previous Game')
fontsheet.write(1, 6, '#')
fontsheet.write(1, 7, 'HOME PLAYER')
fontsheet.write(1, 8, 'Career')
fontsheet.write(1, 9, 'This Season')
fontsheet.write(1, 10, 'Previous Game')
fontsheet.write_merge(18, 18, 0, 11, 'IN-GAME')
fontsheet.write(19, 0, '#')
fontsheet.write(19, 1, 'AWAY PLAYER')
fontsheet.write(19, 2, 'Career')
fontsheet.write(19, 3, 'This Season')
fontsheet.write(19, 4, 'Previous Game')
fontsheet.write(19, 6, '#')
fontsheet.write(19, 7, 'HOME PLAYER')
fontsheet.write(19, 8, 'Career')
fontsheet.write(19, 9, 'This Season')
fontsheet.write(19, 10, 'Previous Game')

away_players = xlrd.open_workbook('Away Stats.xls')
home_players = xlrd.open_workbook('Home Stats.xls')
leaders_list = xlrd.open_workbook('Game Notes.xls')
leaders_list_sheet = leaders_list.sheet_by_name('Leaders Notes')
away_players_career = away_players.sheet_by_name('Last Season')
away_players_this = away_players.sheet_by_name('This Season')
away_players_previous = away_players.sheet_by_name('Previous Game')
home_players_career = home_players.sheet_by_name('Last Season')
home_players_this = home_players.sheet_by_name('This Season')
home_players_previous = home_players.sheet_by_name('Previous Game')
leaderrow = 0
for leaderlistrow in range(leaders_list_sheet.nrows):
    rowdata = leaders_list_sheet.row(leaderlistrow)
    leaders.write(leaderrow, 0, rowdata[0].value)
    leaderrow += 1
away_this = 2
for playerrow in range(away_players_this.nrows):
    rowdata = away_players_this.row(playerrow)
    if rowdata[1].value != '':
        playername = str(rowdata[1].value)
        if playername != 'Full Name':
            playernum = str(int(rowdata[0].value))
            if playernum == '0':
                    extra = 50
            elif playernum == '00':
                    extra = 51
            else:
                extra = 0
            player_details = []
            for value in rowdata:
                if value.value != '':
                    player_details.append(value)
            if len(player_details) - 3 == 6:
                thisseason = 8100 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 8:
                thisseason = 8300 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 10:
                thisseason = 8500 + int(rowdata[0].value) + extra
            else:
                thisseason = 0
            playernum = str(int(rowdata[0].value))
            fontsheet.write(away_this, 0, playernum)
            fontsheet.write(away_this, 1, playername)
            if thisseason != 0:
                fontsheet.write(away_this, 3, thisseason)
            fontsheet.write(away_this + 18, 0, playernum)
            fontsheet.write(away_this + 18, 1, playername)
            if thisseason != 0:
                fontsheet.write(away_this + 18, 3, thisseason + 10000)
            away_this += 1
away_this = 2
for playerrow in range(away_players_previous.nrows):
    rowdata = away_players_previous.row(playerrow)
    if rowdata[1].value != '':
        playername = str(rowdata[1].value)
        if playername != 'Full Name':
            playernum = str(int(rowdata[0].value))
            if playernum == '0':
                    extra = 50
            elif playernum == '00':
                    extra = 51
            else:
                extra = 0
            player_details = []
            for value in rowdata:
                if value.value != '':
                    player_details.append(value)
            if len(player_details) - 3 == 6:
                thisseason = 9100 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 8:
                thisseason = 9300 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 10:
                thisseason = 9500 + int(rowdata[0].value) + extra
            else:
                thisseason = 0
            playernum = str(int(rowdata[0].value))
            if thisseason != 0:
                fontsheet.write(away_this, 4, thisseason)
            if thisseason != 0:
                fontsheet.write(away_this + 18, 4, thisseason + 10000)
            away_this += 1
away_career = 2
for playerrow in range(away_players_career.nrows):
    rowdata = away_players_career.row(playerrow)
    if rowdata[1].value != '':
        playername = str(rowdata[1].value)
        if playername != 'Full Name':
            playernum = str(int(rowdata[0].value))
            if playernum == '0':
                    extra = 50
            elif playernum == '00':
                    extra = 51
            else:
                extra = 0
            player_details = []
            for value in rowdata:
                if value.value != '':
                    player_details.append(value)
            if len(player_details) - 3 == 6:
                thisseason = 7100 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 8:
                thisseason = 7300 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 10:
                thisseason = 7500 + int(rowdata[0].value) + extra
            else:
                thisseason = 0
            playernum = str(int(rowdata[0].value))
            if thisseason != 0:
                fontsheet.write(away_career, 2, thisseason)
            if thisseason != 0:
                fontsheet.write(away_career + 18, 2, thisseason + 10000)
            away_career += 1

home_this = 2
for playerrow in range(home_players_this.nrows):
    rowdata = home_players_this.row(playerrow)
    if rowdata[1].value != '':
        playername = str(rowdata[1].value)
        if playername != 'Full Name':
            playernum = str(int(rowdata[0].value))
            if playernum == '0':
                    extra = 51
            elif playernum == '00':
                    extra = 50
            else:
                extra = 0
            player_details = []
            for value in rowdata:
                if value.value != '':
                    player_details.append(value)
            if len(player_details) - 3 == 6:
                thisseason = 8200 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 8:
                thisseason = 8400 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 10:
                thisseason = 8600 + int(rowdata[0].value) + extra
            else:
                thisseason = 0
            playernum = str(int(rowdata[0].value))
            fontsheet.write(home_this, 6, playernum)
            fontsheet.write(home_this, 7, playername)
            if thisseason != 0:
                fontsheet.write(home_this, 9, thisseason)
            fontsheet.write(home_this + 18, 6, playernum)
            fontsheet.write(home_this + 18, 7, playername)
            if thisseason != 0:
                fontsheet.write(home_this + 18, 9, thisseason + 10000)
            home_this += 1
home_prev = 2
for playerrow in range(home_players_previous.nrows):
    rowdata = home_players_previous.row(playerrow)
    if rowdata[1].value != '':
        playername = str(rowdata[1].value)
        if playername != 'Full Name':
            playernum = str(int(rowdata[0].value))
            if playernum == '0':
                    extra = 51
            elif playernum == '00':
                    extra = 50
            else:
                extra = 0
            player_details = []
            for value in rowdata:
                if value.value != '':
                    player_details.append(value)
            if len(player_details) - 3 == 6:
                thisseason = 9200 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 8:
                thisseason = 9400 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 10:
                thisseason = 9600 + int(rowdata[0].value) + extra
            else:
                thisseason = 0
            playernum = str(int(rowdata[0].value))
            if thisseason != 0:
                fontsheet.write(home_prev, 10, thisseason)
            if thisseason != 0:
                fontsheet.write(home_prev + 18, 10, thisseason + 10000)
            home_prev += 1
home_career = 2
for playerrow in range(home_players_career.nrows):
    rowdata = away_players_career.row(playerrow)
    if rowdata[1].value != '':
        playername = str(rowdata[1].value)
        if playername != 'Full Name':
            playernum = str(int(rowdata[0].value))
            if playernum == '0':
                    extra = 50
            elif playernum == '00':
                    extra = 51
            else:
                extra = 0
            player_details = []
            for value in rowdata:
                if value.value != '':
                    player_details.append(value)
            if len(player_details) - 3 == 6:
                thisseason = 7200 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 8:
                thisseason = 7400 + int(rowdata[0].value) + extra
            elif len(player_details) - 3 == 10:
                thisseason = 7600 + int(rowdata[0].value) + extra
            else:
                thisseason = 0
            playernum = str(int(rowdata[0].value))
            if thisseason != 0:
                fontsheet.write(home_career, 8, thisseason)
            if thisseason != 0:
                fontsheet.write(home_career + 18, 8, thisseason + 10000)
            home_career += 1


# Create This Season Sheet
workbook.save('Font Sheet.xls')

# and print out the html for first game
