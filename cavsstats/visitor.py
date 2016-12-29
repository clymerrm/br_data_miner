import xlwt
from selenium import webdriver

visitor = 'BOS'
visitor_last_game = '201612270BOS'
last_game_qualifier = 'TUE VS MEM'
players = []
player_names_split = []
firstname_whitelist = ['James Michael']

# open phantom js session
browser = webdriver.PhantomJS(service_args=['--ignore-ssl-errors=true'])
browser.maximize_window()

# navigate to team site and parse data
browser.get('http://www.basketball-reference.com/teams/' + visitor + '/2017.html')
player_numbers = browser.find_elements_by_xpath("//*[@id='roster']/tbody/tr/th")
player_numbers = [number.text for number in player_numbers]
player_names = browser.find_elements_by_xpath("//*[@id='roster']/tbody/tr/td[1]")
player_names = [name.text for name in player_names]
player_career = browser.find_elements_by_xpath("//*[@id='roster']/tbody/tr/td[1]/a")
player_career = [career.get_attribute('href') for career in player_career]
player_career = dict(zip(player_names, player_career))
# create roster sheet
for idx, player_name in enumerate(player_names):
    player = {'first':'', 'last':'', 'number':''}
    for whitelist_name in firstname_whitelist:
        if whitelist_name in player_name:
            player['first'] = whitelist_name
            split_name = player_name.replace(whitelist_name, '')
        split_name = player_name.split(' ', 1)
    if player['first'] == '':
        player['first'] = split_name[0]
    player['last'] = " ".join(split_name[1:])
    player['number'] = player_numbers[idx]
    players.append(player)

# create per game stats
per_game_stats = []
per_game_players = browser.find_elements_by_xpath("//*[@id='per_game']/tbody/tr")
for idx, player in enumerate(per_game_players):
    playerstats = {}
    playerstats['name'] = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td/a').text
    playerstats['PTS'] = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="pts_per_g"]').text
    playerstats['FG %'] = str(round(float(browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="fg_pct"]').text)*100, 2))
    threeptfg = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="fg3_pct"]').text
    if threeptfg != '':
        if float(threeptfg) > 0.4:
            playerstats['3-PT FG %'] = str(round(float(threeptfg) * 100, 2))
    rebs = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="trb_per_g"]').text
    if float(rebs) > 1:
        playerstats['REBS'] = rebs
    asts = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="ast_per_g"]').text
    if float(asts) > 1:
        playerstats['ASTS'] = asts
    blks = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="blk_per_g"]').text
    if float(blks) > 1:
        playerstats['BLKS'] = blks
    playerstats['MINS'] = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="mp_per_g"]').text
    oreb = browser.find_element_by_xpath('//*[@id="per_game"]/tbody/tr[' + str(idx + 1) + ']/td[@data-stat="orb_per_g"]').text
    if float(oreb) > 3:
        playerstats['OFF REBS'] = oreb
    season_total_players = browser.find_elements_by_xpath("//*[@id='totals']/tbody/tr/td[1]/a")
    season_total_players = [player.text for player in season_total_players]
    player_index = season_total_players.index(playerstats['name']) + 1
    games_played = browser.find_element_by_xpath('//*[@id="totals"]/tbody/tr[' + str(player_index) + ']/td[@data-stat="g"]/a').text
    steals = browser.find_element_by_xpath('//*[@id="totals"]/tbody/tr[' + str(player_index) + ']/td[@data-stat="stl"]').text
    if int(steals) / int(games_played) > 1:
        playerstats['STLS'] = str(round(int(steals) / int(games_played), 2))
    categories = {'player':playerstats['name'], 'category1':'', 'stats1':'', 'category2':'', 'stats2':'', 'category3':'', 'stats3':'', 'category4':'', 'stats4':'', 'category5':'', 'stats5':'', 'qualifier': 'THIS SEASON'}
    stats_rank = ['FG %', '3-PT FG %', 'OFF REBS', 'REBS', 'ASTS', 'STLS', 'BLKS']
    has_stats = []
    stats = playerstats.keys()
    for stat in stats_rank:
            if stat in stats:
                has_stats.append(stat)
    if playerstats['PTS'] != '0' and len(has_stats) >= 3:
        categories['category1'] = 'PTS'
        categories['stats1'] = playerstats['PTS']
        categories['category2'] = has_stats[0]
        categories['category3'] = has_stats[1]
        categories['category4'] = has_stats[2]
        categories['stats2'] = playerstats[has_stats[0]]
        categories['stats3'] = playerstats[has_stats[1]]
        categories['stats4'] = playerstats[has_stats[2]]
        if len(has_stats) >= 4:
            categories['category5'] = has_stats[3]
            categories['stats5'] = playerstats[has_stats[3]]
    elif playerstats['PTS'] != '0' and len(has_stats) == 2:
        categories['category1'] = 'MINS'
        categories['stats1'] = playerstats['MINS']
        categories['category2'] = 'PTS'
        categories['stats2'] = playerstats['PTS']
        categories['category3'] = has_stats[0]
        categories['stats3'] = playerstats[has_stats[0]]
        categories['category4'] = has_stats[1]
        categories['stats4'] = playerstats[has_stats[1]]
    elif playerstats['PTS'] != '0' and len(has_stats) == 1:
        categories['category1'] = 'MINS'
        categories['stats1'] = playerstats['MINS']
        categories['category3'] = 'PTS'
        categories['stats3'] = playerstats['PTS']
        categories['category5'] = has_stats[0]
        categories['stats5'] = playerstats[has_stats[0]]
    per_game_stats.append(categories)

# create career stats
career_stats = []
for player, url in player_career.items():
    browser.get(url)
    playerstats = {}
    years = browser.find_elements_by_xpath('//*[@id="per_game"]/tbody/tr/th/a')
    years = [year.text for year in years]
    years = str(len(set(years)))
    if years == '1':
        seasons = '(' + years + ' SEASON)'
    else:
        seasons = '(' + years + ' SEASONS)'
    playerstats['name'] = player
    playerstats['PTS'] = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="pts_per_g"]').text
    playerstats['FG %'] = str(round(float(browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="fg_pct"]').text)*100, 2))
    threeptfg = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="fg3_pct"]').text
    if threeptfg != '':
        if float(threeptfg) > 0.4:
            playerstats['3-PT FG %'] = str(round(float(threeptfg) * 100, 2))
    rebs = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="trb_per_g"]').text
    if float(rebs) > 1:
        playerstats['REBS'] = rebs
    asts = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="ast_per_g"]').text
    if float(asts) > 1:
        playerstats['ASTS'] = asts
    blks = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="blk_per_g"]').text
    if float(blks) > 1:
        playerstats['BLKS'] = blks
    playerstats['MINS'] = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="mp_per_g"]').text
    oreb = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="orb_per_g"]').text
    if float(oreb) > 3:
        playerstats['OFF REBS'] = oreb
    games_played = browser.find_element_by_xpath('//*[@id="per_game"]/tfoot/tr[1]/td[@data-stat="g"]').text
    steals = browser.find_element_by_xpath('//*[@id="totals"]/tfoot/tr[1]/td[@data-stat="stl"]').text
    if int(steals) / int(games_played) > 1:
        playerstats['STLS'] = str(round(int(steals) / int(games_played), 2))
    categories = {'player':playerstats['name'], 'category1':'', 'stats1':'', 'category2':'', 'stats2':'', 'category3':'', 'stats3':'', 'category4':'', 'stats4':'', 'category5':'', 'stats5':'', 'qualifier': 'CAREER ' + seasons}
    stats_rank = ['FG %', '3-PT FG %', 'OFF REBS', 'REBS', 'ASTS', 'STLS', 'BLKS']
    has_stats = []
    stats = playerstats.keys()
    for stat in stats_rank:
            if stat in stats:
                has_stats.append(stat)
    if playerstats['PTS'] != '0' and len(has_stats) >= 3:
        categories['category1'] = 'PTS'
        categories['stats1'] = playerstats['PTS']
        categories['category2'] = has_stats[0]
        categories['category3'] = has_stats[1]
        categories['category4'] = has_stats[2]
        categories['stats2'] = playerstats[has_stats[0]]
        categories['stats3'] = playerstats[has_stats[1]]
        categories['stats4'] = playerstats[has_stats[2]]
        if len(has_stats) >= 4:
            categories['category5'] = has_stats[3]
            categories['stats5'] = playerstats[has_stats[3]]
    elif playerstats['PTS'] != '0' and len(has_stats) == 2:
        categories['category1'] = 'MINS'
        categories['stats1'] = playerstats['MINS']
        categories['category2'] = 'PTS'
        categories['stats2'] = playerstats['PTS']
        categories['category3'] = has_stats[0]
        categories['stats3'] = playerstats[has_stats[0]]
        categories['category4'] = has_stats[1]
        categories['stats4'] = playerstats[has_stats[1]]
    elif playerstats['PTS'] != '0' and len(has_stats) == 1:
        categories['category1'] = 'MINS'
        categories['stats1'] = playerstats['MINS']
        categories['category3'] = 'PTS'
        categories['stats3'] = playerstats['PTS']
        categories['category5'] = has_stats[0]
        categories['stats5'] = playerstats[has_stats[0]]
    career_stats.append(categories)

table_id = 'box_' + visitor.lower() + '_basic'

# Create previous game stats
browser.get('http://www.basketball-reference.com/boxscores/' + visitor_last_game + '.html')
browser.find_element_by_xpath('//*[@id="' + table_id + '"]/thead/tr[2]/th[3]').click()
last_game_stats = []
last_game_players = browser.find_elements_by_xpath('//*[@id="' + table_id + '"]/tbody/tr/th/a')
last_game_players = [player.text for player in last_game_players]
for idx, player in enumerate(last_game_players):
    categories_stats = {}
    try:
        categories_stats['MINS'] = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[1]').text
        categories_stats['PTS'] = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[19]').text
        fgsm = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[2]').text
        fgsa = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[3]').text
        if int(fgsm) > 0 and int(fgsm)/int(fgsa) > .10:
            categories_stats['FGS'] = fgsm + '/' + fgsa
        threeptfgsm = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[5]').text
        threeptfgsa = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[6]').text
        if int(threeptfgsm) > 0 and int(threeptfgsm)/int(threeptfgsa) > .40:
            categories_stats['3-PT FGS'] = threeptfgsm + ' / ' + threeptfgsa
        ftsm = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[8]').text
        ftsa = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[9]').text
        if int(ftsm) > 0 and int(ftsm)/int(ftsa) > .9 and int(ftsm) > 10:
            categories_stats['FREE THROWS'] = ftsm + ' / ' + ftsa
        oreb = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[11]').text
        if int(oreb) > 3:
            categories_stats['OFF REB'] = oreb
        rebs = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[13]').text
        if int(rebs) > 3:
            categories_stats['REBS'] = rebs
        asts = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[14]').text
        if int(asts) > 1:
            categories_stats['ASTS'] = asts
        stls = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[15]').text
        if int(stls) > 1:
            categories_stats['STLS'] = stls
        blks = browser.find_element_by_xpath('//*[@id="' + table_id + '"]/tbody/tr[' + str(idx + 1) + ']/td[16]').text
        if int(blks) > 1:
            categories_stats['BLKS'] = blks
        categories = {'player':player, 'category1':'', 'stats1':'', 'category2':'', 'stats2':'', 'category3':'', 'stats3':'', 'category4':'', 'stats4':'', 'category5':'', 'stats5':'', 'qualifier': 'This Season'}
        stats_rank = ['FGS', '3-PT FGS', 'FREE THROWS', 'OFF REB', 'REBS', 'ASTS', 'STLS', 'BLKS']
        has_stats = []
        stats = categories_stats.keys()
        for stat in stats_rank:
                if stat in stats:
                    has_stats.append(stat)
        if categories_stats['PTS'] != '0' and len(has_stats) >= 3:
            categories['category1'] = 'PTS'
            categories['stats1'] = categories_stats['PTS']
            categories['category2'] = has_stats[0]
            categories['category3'] = has_stats[1]
            categories['category4'] = has_stats[2]
            categories['stats2'] = categories_stats[has_stats[0]]
            categories['stats3'] = categories_stats[has_stats[1]]
            categories['stats4'] = categories_stats[has_stats[2]]
            if len(has_stats) >= 4:
                categories['category5'] = has_stats[3]
                categories['stats5'] = categories_stats[has_stats[3]]
        elif categories_stats['PTS'] != '0' and len(has_stats) == 2:
            categories['category1'] = 'PTS'
            categories['stats1'] = categories_stats['PTS']
            categories['category3'] = has_stats[0]
            categories['stats3'] = categories_stats[has_stats[0]]
            categories['category5'] = has_stats[1]
            categories['stats5'] = categories_stats[has_stats[1]]
        elif categories_stats['PTS'] != '0' and len(has_stats) == 1:
            categories['category1'] = 'MINS'
            categories['stats1'] = categories_stats['MINS']
            categories['category3'] = 'PTS'
            categories['stats3'] = categories_stats['PTS']
            categories['category5'] = has_stats[0]
            categories['stats5'] = categories_stats[has_stats[0]]
        last_game_stats.append(categories)
    except:
        pass

# Create First and Last Name Workbook
workbook = xlwt.Workbook()
lastseason = workbook.add_sheet('Last Season')
thisseason = workbook.add_sheet('This Season')
previousgame = workbook.add_sheet('Previous Game')
firstname = workbook.add_sheet('First & Last Name')
firstname.write(0, 0, 'No')
firstname.write(0, 1, 'First Name')
firstname.write(0, 2, 'Last Name')
lastseason.write(0, 0, 'No')
lastseason.write(0, 1, 'Full Name')
lastseason.write(0, 2, 'Category1')
lastseason.write(0, 3, 'Stat1')
lastseason.write(0, 4, 'Category2')
lastseason.write(0, 5, 'Stat2')
lastseason.write(0, 6, 'Category3')
lastseason.write(0, 7, 'Stat3')
lastseason.write(0, 8, 'Category4')
lastseason.write(0, 9, 'Stat4')
lastseason.write(0, 10, 'Category5')
lastseason.write(0, 11, 'Stat5')
lastseason.write(0, 12, 'Qualifier')
thisseason.write(0, 0, 'No')
thisseason.write(0, 1, 'Full Name')
thisseason.write(0, 2, 'Category1')
thisseason.write(0, 3, 'Stat1')
thisseason.write(0, 4, 'Category2')
thisseason.write(0, 5, 'Stat2')
thisseason.write(0, 6, 'Category3')
thisseason.write(0, 7, 'Stat3')
thisseason.write(0, 8, 'Category4')
thisseason.write(0, 9, 'Stat4')
thisseason.write(0, 10, 'Category5')
thisseason.write(0, 11, 'Stat5')
thisseason.write(0, 12, 'Qualifier')
previousgame.write(0, 0, 'No')
previousgame.write(0, 1, 'Full Name')
previousgame.write(0, 2, 'Category1')
previousgame.write(0, 3, 'Stat1')
previousgame.write(0, 4, 'Category2')
previousgame.write(0, 5, 'Stat2')
previousgame.write(0, 6, 'Category3')
previousgame.write(0, 7, 'Stat3')
previousgame.write(0, 8, 'Category4')
previousgame.write(0, 9, 'Stat4')
previousgame.write(0, 10, 'Category5')
previousgame.write(0, 11, 'Stat5')
previousgame.write(0, 12, 'Qualifier')
row = 1
col = 0
for x in range(1, 100):
    if x == 50:
        x = '00'
    elif x == 51:
        x = '0'
    firstname.write(row, col, x)
    lastseason.write(row, col, x)
    thisseason.write(row, col, x)
    previousgame.write(row, col, x)
    row += 1
for player in players:
    x = int(player['number'])
    if player['number'] == '00':
        x = 50
    elif player['number'] == '0':
        x = 51
    full_name = player['first'] + ' ' + player['last']
    firstname.write(x, col + 1, player['first'].upper())
    firstname.write(x, col + 2, player['last'].upper())
    lastseason.write(x, col + 1, full_name.upper())
    lastseason.write(x, 13, xlwt.Formula('IF(COUNTIF(C' + str(x + 1) + ':L' + str(x + 1) + ',"*")>0,COUNTIF(C' + str(x + 1) + ':L' + str(x + 1) + ',"*"),"")'))
    lastseason.write(x, 14, xlwt.Formula('IF(N' + str(x + 1) + '=6,7200,IF(N' + str(x + 1) + '=8,7400,IF(N' + str(x + 1) + '=10,7600,0)))'))
    thisseason.write(x, col + 1, full_name.upper())
    thisseason.write(x, 13, xlwt.Formula('IF(COUNTIF(C' + str(x + 1) + ':L' + str(x + 1) + ',"*")>0,COUNTIF(C' + str(x + 1) + ':L' + str(x + 1) + ',"*"),"")'))
    thisseason.write(x, 14, xlwt.Formula('IF(N' + str(x + 1) + '=6,8200,IF(N' + str(x + 1) + '=8,8400,IF(N' + str(x + 1) + '=10,8600,0)))'))
    previousgame.write(x, col + 1, full_name.upper())
    previousgame.write(x, 13, xlwt.Formula('IF(COUNTIF(C' + str(x + 1) + ':L' + str(x + 1) + ',"*")>0,COUNTIF(C' + str(x + 1) + ':L' + str(x + 1) + ',"*"),"")'))
    previousgame.write(x, 14, xlwt.Formula('IF(N' + str(x + 1) + '=6,9200,IF(N' + str(x + 1) + '=8,9400,IF(N' + str(x + 1) + '=10,9600,0)))'))
    player_per_game_stats = []
    for idx, player in enumerate(career_stats):
        if player['player'] == full_name:
            player_stats_index = idx
            career_per_game_stats = career_stats[player_stats_index]
            lastseason.write(x, col + 2, career_per_game_stats['category1'])
            lastseason.write(x, col + 3, career_per_game_stats['stats1'])
            lastseason.write(x, col + 4, career_per_game_stats['category2'])
            lastseason.write(x, col + 5, career_per_game_stats['stats2'])
            lastseason.write(x, col + 6, career_per_game_stats['category3'])
            lastseason.write(x, col + 7, career_per_game_stats['stats3'])
            lastseason.write(x, col + 8, career_per_game_stats['category4'])
            lastseason.write(x, col + 9, career_per_game_stats['stats4'])
            lastseason.write(x, col + 10, career_per_game_stats['category5'])
            lastseason.write(x, col + 11, career_per_game_stats['stats5'])
            lastseason.write(x, col + 12, career_per_game_stats['qualifier'])
    for idx, player in enumerate(per_game_stats):
        if player['player'] == full_name:
            player_stats_index = idx
            player_per_game_stats = per_game_stats[player_stats_index]
            thisseason.write(x, col + 2, player_per_game_stats['category1'])
            thisseason.write(x, col + 3, player_per_game_stats['stats1'])
            thisseason.write(x, col + 4, player_per_game_stats['category2'])
            thisseason.write(x, col + 5, player_per_game_stats['stats2'])
            thisseason.write(x, col + 6, player_per_game_stats['category3'])
            thisseason.write(x, col + 7, player_per_game_stats['stats3'])
            thisseason.write(x, col + 8, player_per_game_stats['category4'])
            thisseason.write(x, col + 9, player_per_game_stats['stats4'])
            thisseason.write(x, col + 10, player_per_game_stats['category5'])
            thisseason.write(x, col + 11, player_per_game_stats['stats5'])
            thisseason.write(x, col + 12, player_per_game_stats['qualifier'])
    for idx, player in enumerate(last_game_stats):
        if player['player'] == full_name:
            player_stats_index = idx
            last_game_player_stats = last_game_stats[player_stats_index]
            previousgame.write(x, col + 2, last_game_player_stats['category1'])
            previousgame.write(x, col + 3, last_game_player_stats['stats1'])
            previousgame.write(x, col + 4, last_game_player_stats['category2'])
            previousgame.write(x, col + 5, last_game_player_stats['stats2'])
            previousgame.write(x, col + 6, last_game_player_stats['category3'])
            previousgame.write(x, col + 7, last_game_player_stats['stats3'])
            previousgame.write(x, col + 8, last_game_player_stats['category4'])
            previousgame.write(x, col + 9, last_game_player_stats['stats4'])
            previousgame.write(x, col + 10, last_game_player_stats['category5'])
            previousgame.write(x, col + 11, last_game_player_stats['stats5'])
            previousgame.write(x, col + 12, last_game_qualifier)

# for player in per_game_stats:
#     x = int(player['name'])

# Create This Season Sheet
workbook.save('Away Stats.xls')

# and print out the html for first game
