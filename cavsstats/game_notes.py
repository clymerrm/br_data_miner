import xlwt
from selenium import webdriver

last_game_qualifier = 'FRI VS MIA'
players = []
player_names_split = []
firstname_whitelist = ['James Michael']

# open phantom js session
browser = webdriver.PhantomJS(service_args=['--ignore-ssl-errors=true'])
browser.maximize_window()

browser.get('http://www.basketball-reference.com/teams/CLE/2017.html')
player_names = browser.find_elements_by_xpath("//*[@id='roster']/tbody/tr/td[1]")
player_names = [name.text for name in player_names]

# navigate to team site and parse data
workbook = xlwt.Workbook()
leaders = workbook.add_sheet('Leaders')
notes = workbook.add_sheet('Leaders Notes')
leaders.write(0, 0, 'Title')
leaders.write(0, 1, 'Subtitle')
leaders.write(0, 2, 'Rank 1')
leaders.write(0, 3, 'Player 1')
leaders.write(0, 4, 'Number 1')
leaders.write(0, 5, 'Rank 2')
leaders.write(0, 6, 'Player 2')
leaders.write(0, 7, 'Number 2')
leaders.write(0, 8, 'Rank 3')
leaders.write(0, 9, 'Player 3')
leaders.write(0, 10, 'Number 3')
leaders.write(0, 11, 'Rank 4')
leaders.write(0, 12, 'Player 4')
leaders.write(0, 13, 'Number 4')
leaders.write(0, 14, 'Rank 5')
leaders.write(0, 15, 'Player 5')
leaders.write(0, 16, 'Number 5')
notes.write(0, 0, 'Milestone Watch')
playernotes = []
row = 1
stats = ['g_career', 'pts_career', 'fg_career', 'fg3_career', 'orb_career', 'drb_career', 'trb_career', 'ast_career', 'stl_career', 'blk_career', 'triple-double-most-times']
heading = ['GAMES PLAYED', 'POINTS', 'FIELD GOALS MADE', '3-PT FIELD GOALS MADE', 'OFFENSIVE REBOUNDS', 'DEFENSIVE REBOUNDS', 'REBOUNDS', 'ASSISTS', 'STEALS', 'BLOCKS', 'TRIPLE DOUBLES']
grouping = ['NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY', 'NBA HISTORY']
for masteridx, stat in enumerate(stats):
    browser.get('http://www.basketball-reference.com/leaders/' + stat + '.html')
    players = browser.find_elements_by_xpath("//*[@id='nba']/tbody/tr/td[2]")
    players = [player.text.replace('*', '') for player in players]
    ranknumbers = browser.find_elements_by_xpath("//*[@id='nba']/tbody/tr/td[1]")
    ranknumbers = [ranknumber.text for ranknumber in ranknumbers]
    playerstats = browser.find_elements_by_xpath("//*[@id='nba']/tbody/tr/td[3]")
    playerstats = [stat.text for stat in playerstats]
    for idx, player in enumerate(players):
        if player in player_names:
            ranks = {'1name':'','1stat':'','2name':'','2stat':'','3name':'','3stat':'','4name':'','4stat':'','5name':'','5stat':'', '1rank':'', '2rank':'', '3rank':'', '4rank':'', '5rank':''}
            ranks['1name'] = players[idx-2]
            ranks['2name'] = players[idx-1]
            ranks['3name'] = players[idx]
            ranks['4name'] = players[idx+1]
            ranks['5name'] = players[idx+2]
            ranks['1stat'] = playerstats[idx-2]
            ranks['2stat'] = playerstats[idx-1]
            ranks['3stat'] = playerstats[idx]
            ranks['4stat'] = playerstats[idx+1]
            ranks['5stat'] = playerstats[idx+2]
            ranks['1rank'] = ranknumbers[idx-2].replace('.','')
            if ranks['1rank'] == ' ':
                ranks['1rank'] = ranknumbers[idx-3].replace('.','')
            if ranks['1rank'] == ' ':
                ranks['1rank'] = ranknumbers[idx-4].replace('.','')
            ranks['2rank'] = ranknumbers[idx-1].replace('.','')
            ranks['3rank'] = ranknumbers[idx].replace('.','')
            ranks['4rank'] = ranknumbers[idx+1].replace('.','')
            ranks['5rank'] = ranknumbers[idx+2].replace('.','')
            if ranks['2rank'] == ' ':
                nextrank = ranks['1rank']
            else:
                nextrank = ranks['2rank']
            notesline = ranks['3name'] + ' is ' + str(int(ranks['2stat']) - int(ranks['3stat']) + 1) + ' ' + heading[masteridx] + ' from moving to ' + nextrank + ' in ' + heading[masteridx]
            playernotes.append(notesline)
            leaders.write(row, 0, grouping[masteridx])
            leaders.write(row, 1, heading[masteridx])
            leaders.write(row, 2, ranks['1rank'])
            leaders.write(row, 3, ranks['1name'].upper())
            leaders.write(row, 4, ranks['1stat'])
            leaders.write(row, 5, ranks['2rank'])
            leaders.write(row, 6, ranks['2name'].upper())
            leaders.write(row, 7, ranks['2stat'])
            leaders.write(row, 8, ranks['3rank'])
            leaders.write(row, 9, ranks['3name'].upper())
            leaders.write(row, 10, ranks['3stat'])
            leaders.write(row, 11, ranks['4rank'])
            leaders.write(row, 12, ranks['4name'].upper())
            leaders.write(row, 13, ranks['4stat'])
            leaders.write(row, 14, ranks['5rank'])
            leaders.write(row, 15, ranks['5name'].upper())
            leaders.write(row, 16, ranks['5stat'])
            row += 1

row = 1
playernotes.sort()
for note in playernotes:
    notes.write(row, 0, note)
    row += 1


# Create This Season Sheet
workbook.save('Game Notes.xls')