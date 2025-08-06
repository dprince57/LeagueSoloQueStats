import requests
import xlwt


#Solo Queue Data Analysis - create a model that can identify and compile solo queue data of NALCS pros (smurfs and mains

wb = xlwt.Workbook()
win = 0.0
print("Welcome to a concept on SoloQ Stats v0.1 by Dakota Prince.")
print("This can disperse more data, only needs to be added.")

def lane(n, w, t):
    if n['lane'] == "TOP":
        w.write(t + 1, 2, "Top")
    elif n['lane'] == "JUNGLE":
        w.write(t + 1, 2, "Jungle")
    elif n['role'] == "DUO_CARRY":
        w.write(t + 1, 2, "ADC")
    elif n['role'] == "DUO_SUPPORT":
        w.write(t + 1, 2, "Support")
    else:
        w.write(t + 1, 2, "Mid")

def getChampData():

    r = requests.get("https://ddragon.leagueoflegends.com/cdn/8.19.1/data/en_US/champion.json")
    return r.json()


def requestSumData(n, a):

    URL = "https://NA1.api.riotgames.com/lol/summoner/v4/summoners/by-name/" + n + "?api_key=" + a
    r = requests.get(URL)
    return r.json()


def requestListData(i, a, num):

    URL = "https://NA1.api.riotgames.com/lol/match/v4/matchlists/by-account/" + i + "?endIndex=" + num + "&api_key=" + a
    r = requests.get(URL)
    return r.json()


def getMatchData(i, a):

    URL = "https://NA1.api.riotgames.com/lol/match/v4/matches/" + str(i) + "?api_key=" + a
    r = requests.get(URL)
    return r.json()


def printMatchStats(ms, t, w):

    #K/D/A (k/d)(a/d)(d/d) Drop the deaths for k/a <- for supports
    #K/D <- for laners and jungle
    global win

    w.write(t + 1, 1, t)
    w.write(t + 1, 4, int(ms['stats']['kills']))
    w.write(t + 1, 5, int(ms['stats']['deaths']))
    w.write(t + 1, 6, int(ms['stats']['assists']))
    w.write(t + 1, 8, int(ms['stats']['visionScore']))
    if str(ms['stats']['win']) == 'True':
        w.write(t + 1, 7, "Win")
        win += 1
    else:
        w.write(t + 1, 7, "Loss")


def main():
    global win
    champ = getChampData()

    name = input('Please type summoner name (WARNING!: IF NOT SPELT CORRECTLY WILL THROW ERROR)\n')
    num = input('how many matches would you like to obtain (1-40)')
    print("Only SoloQ games will be calculated.")
    #change API Key daily or it will not work
    api_key = 'KEY HERE'

    ws = wb.add_sheet(str(name))
    ws.write(0, 0, name)
    ws.write(0, 1, "Match #")
    ws.write(0, 2, "Role")
    ws.write(0, 3, "Champion")
    ws.write(0, 4, "Kills")
    ws.write(0, 5, "Deaths")
    ws.write(0, 6, "Assists")
    ws.write(0, 7, "W/L")
    ws.write(0, 8, "Vision Score")

    ws.write(0, 13, "KD")
    ws.write(0, 14, "W/L %")
    ws.write(0, 15, "Avg Vision")
    ws.write(1, 13, xlwt.Formula("AVERAGE(E3:E43,F3:F43)"))
    ws.write(1, 15, xlwt.Formula("AVERAGE(I3:I43)"))



    r = requestSumData(name, api_key)


    ID = str(r['accountId'])

    r2 = requestListData(ID, api_key, num)

    total = 0
    print("Writing Stats for " + name)
    #Nested For Loops to parse the 4 API Pages needed to find match data
    for x in r2['matches']:
        if x['queue'] == 420:
            total += 1
            lane(x, ws, total)
            for ch in champ['data']:
                if str(champ['data'][ch]['key']) == str(x['champion']):
                    ws.write(total + 1, 3, str(ch))
                    break

            r3 = getMatchData(x['gameId'], api_key)

            for t in r3['participantIdentities']:

                if t['player']['accountId'] == ID:
                    tempI = t['participantId']

                    for s in r3['participants']:
                        if s['participantId'] == tempI:
                            printMatchStats(s, total, ws)

    ws.write(1, 14, "%.0f%%" % (100 * win/int(num)))
    y = ''
    while y != 'yes':
        y = input("Writing complete. Would you like to select another summoner? [yes/no]\n")
        if y == 'yes':
            win = 0.0
            main()
        if y == 'no':
            print(".xls file saved to root directory. Good Day Summoner!")
            break


if __name__ == '__main__':
    main()

wb.save('SoloQStats.xls')
