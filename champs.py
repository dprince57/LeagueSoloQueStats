import requests
import xlwt
import time

wb = xlwt.Workbook()
champs = wb.add_sheet('champs', cell_overwrite_ok=True)
items = wb.add_sheet('items', cell_overwrite_ok=True)
print("Updating Excel files now for new patch!")


def addItems():
    global items
    i = requests.get("http://ddragon.leagueoflegends.com/cdn/6.24.1/data/en_US/item.json")
    c = i.json()
    t = 0
    items.write(0, 0, 'Name')
    items.write(1, 0, 'FlatHpPoolMod')
    items.write(2, 0, 'FlatMPPoolMod')
    items.write(3, 0, 'PercentHPPoolMod')
    items.write(4, 0, 'PercentMPPoolMod')
    items.write(5, 0, 'FlatHPRegenMod')
    items.write(6, 0, 'PercentHPRegenMod')
    items.write(7, 0, 'FlatMPRegenMod')
    items.write(8, 0, 'PercentMPRegenMod')
    items.write(9, 0, 'FlatArmorMod')
    items.write(10, 0, 'PercentArmorMod')
    items.write(11, 0, 'FlatPhysicalDamageMod')
    items.write(12, 0, 'PercentPhysicalDamageMod')
    items.write(13, 0, 'FlatMagicDamageMod')
    items.write(14, 0, 'PercentMagicDamageMod')
    items.write(15, 0, 'FlatMovementSpeedMod')
    items.write(16, 0, 'PercentMovementSpeedMod')
    items.write(17, 0, 'FlatAttackSpeedMod')
    items.write(18, 0, 'PercentAttackSpeedMod')
    items.write(19, 0, 'FlatCritChanceMod')
    items.write(20, 0, 'PercentCritChanceMod')
    items.write(21, 0, 'Cost')

    for x in c['data']:
        items.write(0, t + 1, c['data'][x]['name'])
        for s in c['data'][x]['stats']:
            if s == 'FlatHpPoolMod':
                items.write(1, t + 1, c['data'][x]['stats']['FlatHpPoolMod'])

            if s =='FlatMPPoolMod':
                items.write(2, t + 1, c['data'][x]['stats']['FlatMPPoolMod'])

            if s =='PercentHPPoolMod':
                items.write(3, t + 1, c['data'][x]['stats']['PercentHPPoolMod'])

            if s =='PercentMPPoolMod':
                items.write(4, t + 1, c['data'][x]['stats']['PercentMPPoolMod'])

            if s =='FlatHPRegenMod':
                items.write(5, t + 1, c['data'][x]['stats']['FlatHPRegenMod'])

            if s == 'PercentHPRegenMod':
                items.write(6, t + 1, c['data'][x]['stats']['PercentHPRegenMod'])

            if s == 'FlatMPRegenMod':
                items.write(7, t + 1, c['data'][x]['stats']['FlatMPRegenMod'])

            if s == 'PercentMPRegenMod':
                items.write(8, t + 1, c['data'][x]['stats']['PercentMPRegenMod'])

            if s == 'FlatArmorMod':
                items.write(9, t + 1, c['data'][x]['stats']['FlatArmorMod'])

            if s == 'PercentArmorMod':
                items.write(10, t + 1, c['data'][x]['stats']['PercentArmorMod'])

            if s == 'FlatPhysicalDamageMod':
                items.write(11, t + 1, c['data'][x]['stats']['FlatPhysicalDamageMod'])

            if s == 'PercentPhysicalDamageMod':
                items.write(12, t + 1, c['data'][x]['stats']['PercentPhysicalDamageMod'])

            if s == 'FlatMagicDamageMod':
                items.write(13, t + 1, c['data'][x]['stats']['FlatMagicDamageMod'])

            if s == 'PercentMagicDamageMod':
                items.write(14, t + 1, c['data'][x]['stats']['PercentMagicDamageMod'])

            if s == 'FlatMovementSpeedMod':
                items.write(15, t + 1, c['data'][x]['stats']['FlatMovementSpeedMod'])

            if s == 'PercentMovementSpeedMod':
                items.write(16, t + 1, c['data'][x]['stats']['PercentMovementSpeedMod'])

            if s == 'FlatAttackSpeedMod':
                items.write(17, t + 1, c['data'][x]['stats']['FlatAttackSpeedMod'])

            if s == 'PercentAttackSpeedMod':
                items.write(18, t + 1, c['data'][x]['stats']['PercentAttackSpeedMod'])

            if s == 'FlatCritChanceMod' :
                items.write(19, t + 1, c['data'][x]['stats']['FlatCritChanceMod'])

            if s == 'PercentCritChanceMod' :
                items.write(20, t + 1, c['data'][x]['stats']['PercentCritChanceMod'])

        items.write(21, t + 1, c['data'][x]['gold']['base'])
        t += 1

def getChamps():

    global champs
    champ = requests.get("https://ddragon.leagueoflegends.com/cdn/8.19.1/data/en_US/champion.json")
    c = champ.json()
    t = 0
    champs.write(0, 0, 'Name')
    champs.write(1, 0, 'hp')
    champs.write(2, 0, 'hpperlevel')
    champs.write(3, 0, 'mp')
    champs.write(4, 0, 'mpperlevel')
    champs.write(5, 0, 'movespeed')
    champs.write(6, 0, 'armor')
    champs.write(7, 0, 'armorperlevel')
    champs.write(8, 0, 'spellblock')
    champs.write(9, 0, 'spellblockperlevel')
    champs.write(10, 0, 'attackrange')
    champs.write(11, 0, 'hpregen')
    champs.write(12, 0, 'hpregenperlevel')
    champs.write(13, 0, 'mpregen')
    champs.write(14, 0, 'mpregenperlevel')
    champs.write(15, 0, 'crit')
    champs.write(16, 0, 'critperlevel')
    champs.write(17, 0, 'attackdamage')
    champs.write(18, 0, 'attackdamageperlevel')
    champs.write(19, 0, 'attackspeedoffset')
    champs.write(20, 0, 'attackspeedperlevel')

    for x in c['data']:
        champs.write(0, t + 1, x)
        champs.write(1, t + 1, c['data'][x]['stats']['hp'])
        champs.write(2, t + 1, c['data'][x]['stats']['hpperlevel'])
        champs.write(3, t + 1, c['data'][x]['stats']['mp'])
        champs.write(4, t + 1, c['data'][x]['stats']['mpperlevel'])
        champs.write(5, t + 1, c['data'][x]['stats']['movespeed'])
        champs.write(6, t + 1, c['data'][x]['stats']['armor'])
        champs.write(7, t + 1, c['data'][x]['stats']['armorperlevel'])
        champs.write(8, t + 1, c['data'][x]['stats']['spellblock'])
        champs.write(9, t + 1, c['data'][x]['stats']['spellblockperlevel'])
        champs.write(10, t + 1, c['data'][x]['stats']['attackrange'])
        champs.write(11, t + 1, c['data'][x]['stats']['hpregen'])
        champs.write(12, t + 1, c['data'][x]['stats']['hpregenperlevel'])
        champs.write(13, t + 1, c['data'][x]['stats']['mpregen'])
        champs.write(14, t + 1, c['data'][x]['stats']['mpregenperlevel'])
        champs.write(15, t + 1, c['data'][x]['stats']['crit'])
        champs.write(16, t + 1, c['data'][x]['stats']['critperlevel'])
        champs.write(17, t + 1, c['data'][x]['stats']['attackdamage'])
        champs.write(18, t + 1, c['data'][x]['stats']['attackdamageperlevel'])
        champs.write(19, t + 1, c['data'][x]['stats']['attackspeedoffset'])
        champs.write(20, t + 1, c['data'][x]['stats']['attackspeedperlevel'])
        t += 1

def main():
    getChamps()
    addItems()


main()
wb.save("updated.xls")
print("Update Complete.")
time.sleep(6)
