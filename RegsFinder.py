import glob, time


def GetSheet(n):  # Get data about spreadsheet
    # get the doc from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    # get the XText interface
    sheet = model.Sheets.getByIndex(n)
    return sheet


def GetPath(*args):  # modify path to directory
    sheet = GetSheet(0)

    if sheet.getCellByPosition(6, 1).String == "":
        sheet.getCellByPosition(6, 1).String = ("Podaj katalog!")
        return
    else:
        path = sheet.getCellByPosition(6, 1).String
        path = path.replace("\\", "\\\\") + "\\\\"
    return path


def ClearTable(*args):  # Clear results
    sheet = GetSheet(0)
    maxw = int(sheet.getCellByPosition(1, 1).Value)
    maxk = int(sheet.getCellByPosition(11, 1).Value)
    for l in range(maxw):
        for m in range(maxk):
            sheet.getCellByPosition(m, l + 2).String = ""
    for i in range(2,7):
        sheet.getCellByPosition(14, i).String = ""
        sheet.getCellByPosition(15, i).String = ""

def PrintList(List, x, y):
    sheet = GetSheet(0)
    for n in range(len(List)):
        for m in range(len(List[n])):
            sheet.getCellByPosition(x + m, n + y).Value = List[n][m]
    sheet.getCellByPosition(1, 1).Value = len(List) + 3
    sheet.getCellByPosition(11, 1).Value = max(len(List) + 2, 15)


def CheckIfInTheList(ID, List):
    # sprawdza czy ID jest już na posortowanej liście i jeżeli jest to zwiększa licznik rozdań.
    # Jeżeli elementu nie ma na liście to wstawia go na właściwe miejsce.
#    sheet = GetSheet(0)
    flaga = 0
#    x = 0
    for p in range(len(List)):
        if List[p][0] == ID:
            flaga = 1
            List[p][1] += 1
#            sheet.getCellByPosition(6, 3 + len(List)).Value = ID
#            sheet.getCellByPosition(7, 3 + len(List)).Value = List[p][1]
#            sheet.getCellByPosition(8, 3 + len(List)).Value = p
#            x += 1
            break
        if List[p][0] > ID:
            flaga = 2
            break
    if flaga == 0:
        List.append([ID, 1])  # Wstawia kolejny element
    if flaga == 2:
        List.append([ID, 1])  # Wstawia kolejny element
        m = len(List)
        if m > 1:
            while List[m - 2][0] > ID:
                List[m - 1][0] = List[m - 2][0]
                List[m - 1][1] = List[m - 2][1]
                m = m - 1
                if m == 1:
                    break
            List[m - 1][0] = ID
            List[m - 1][1] = 1


def RegsFinder(*args):
    # Path to tree
    path = GetPath()
    sheet = GetSheet(0)
    minhands = sheet.getCellByPosition(5, 1).Value
    IdToCheck = sheet.getCellByPosition(10, 1).Value
    start = time.time()
    timer = time.time()
    dircounter = 0
    directories = [d for d in glob.glob(path + "*\\", recursive=True)]
    Players = []
    PlayersByHands = []
    Hands = []
    hands = 0
    try:
        with open(path + "\\hands.txt", encoding="utf8") as handsfile:
            i = 0
            lines = handsfile.readlines()
            while i < len(lines):
                handtemp = []
                j = 0
                while j < len(lines[i].split()):
                    handtemp.append(int(lines[i].split()[j]))
                    j += 1
                Hands.append(handtemp)
                hands +=1
                i += 1
    except:
        with open(path + "\\hands.txt", "w", encoding="utf8") as handsfile:
            for d in directories:
                dircounter += 1
                files = [f for f in glob.glob(d + "**/*.txt", recursive=True)]
        #       sessions = 0
                for f in files:
                    with open(f, encoding="utf8") as plik:
        #                   sessions += 1  # licznik sesji
                        #                if sheet.getCellByPosition(5, 1).Value == 0:
                        #                    sheet.getCellByPosition(6, sessions+2).String = f
                        i = 0
                        lines = plik.readlines()
                        while i < len(lines):
                            if lines[i].strip().count("PokerStars") == 1:
                                #                        sheet.getCellByPosition(0, 3 + hands).Value = i
                                #                        sheet.getCellByPosition(1, 3 + hands).String = lines[i]
                                handtemp = []
                                j = i + 2
                                while j <= len(lines):
                                    if lines[j].split()[0] == "Seat":
                                        #                                sheet.getCellByPosition(3 + j - i - 2, 3 + hands).Value = int(lines[j].split()[2])
                                        #                                sheet.getCellByPosition(3 + j - i - 2, 3 + hands).Value = j
                                        handtemp.append(int(lines[j].split()[2]))
                                        handsfile.write(lines[j].split()[2] + " ")
        #                               handsfile.write(" ")
                                        j += 1
                                    else:
                                        break
                                Hands.append(handtemp)
                                handsfile.write("\n")
                                hands += 1  # licznik rozdań
#                           if hands % 1000 == 0:
#                               sheet.getCellByPosition(12, 1 + hands // 1000).Value = time.time() - timer
#                               timer = time.time()
                                i = j
                            else:
                                i += 1

    sheet.getCellByPosition(14, 2).String = "Import plików"
    sheet.getCellByPosition(15, 2).Value = time.time() - start
    timer = time.time()

    sheet3 = GetSheet(2)
    IDs = []
    r = 0
    while sheet3.getCellByPosition(0, r).Value != 0:
        IDs.append(sheet3.getCellByPosition(0, r).Value)
        r += 1

    try:
        with open(path + "\\players.txt", encoding="utf8") as playersfile:
            i = 0
            lines = playersfile.readlines()
            while i < len(lines):
                Players.append([int(lines[i].split()[0]),int(lines[i].split()[1])])
                i += 1
            sheet.getCellByPosition(14, 3).String = "Tworzenie listy graczy:"
            sheet.getCellByPosition(15, 3).Value = time.time() - timer
            timer = time.time()
            PlayersByHands = sorted(Players, reverse=True, key=lambda list: (list[1], list[0]))
            sheet.getCellByPosition(14, 4).String = "Sortowanie listy graczy:"
            sheet.getCellByPosition(15, 4).Value = time.time() - timer
            timer = time.time()
    except:
        for h in range(len(Hands)):
            for j in range(len(Hands[h])):
                CheckIfInTheList(Hands[h][j], Players)
#            sheet.getCellByPosition(10, 3 + j).Value = Hands[h][j]
        sheet.getCellByPosition(14, 3).String = "Tworzenie listy graczy:"
        sheet.getCellByPosition(15, 3).Value = time.time() - timer
        timer = time.time()
        PlayersByHands = sorted(Players, reverse=True, key=lambda list: (list[1], list[0]))
        sheet.getCellByPosition(14, 4).String = "Sortowanie listy graczy:"
        sheet.getCellByPosition(15, 4).Value = time.time() - timer
        timer = time.time()
        with open(path + "\\players.txt", "w", encoding="utf8") as playersfile:
            for player in Players:
                playersfile.write(str(player[0]) + " " + str(player[1]) + "\n")

    if minhands == 0:
        PrintList(Players, 0, 3)
        PrintList(PlayersByHands, 3, 3)
#        PrintList(Hands,3,3)
        sheet.getCellByPosition(4, 1).Value = dircounter
        sheet.getCellByPosition(2, 1).Value = hands
        sheet.getCellByPosition(3, 1).Value = len(PlayersByHands)
        sheet.getCellByPosition(1, 1).Value = len(PlayersByHands) + 10
        sheet.getCellByPosition(11, 1).Value = 5
        sheet.getCellByPosition(14, 5).String = "Drukowanie listy graczy:"
        sheet.getCellByPosition(15, 5).Value = time.time() - timer
        timer = time.time()


    if minhands > 0:
        Relations = []
        sheet2 = GetSheet(1)
        KnownIDs = []
        r = 0
        while sheet2.getCellByPosition(0, r).Value != 0:
            KnownIDs.append(sheet2.getCellByPosition(0, r).Value)
            r +=1
        n = 0
        while PlayersByHands[n][1] > minhands:
            Known = False
            for ID in KnownIDs:
                if PlayersByHands[n][0] == ID:
#                    sheet.getCellByPosition(n, 30).Value = ID
                    Known = True
                    break
            if Known == False:
                Relations.append([PlayersByHands[n][0], PlayersByHands[n][1]])
            n += 1
        Relations = sorted(Relations, key=lambda list: (list[0], list[1]))
        for item in Relations:
            for n in range(len(Relations)):
                item.append(0)
        for hand in Hands:
            for i in range(len(hand)):
                if hand[i] > Relations[len(Relations) - 1][0]:
                    hand[i] = 0
                else:
                    for j in range(len(Relations)):
                        if Relations[j][0] == hand[i]:
                            hand[i] = j + 1
                            break
                        if Relations[j][0] > hand[i]:
                            hand[i] = 0
                            break
        for hand in Hands:
            for i in range(len(hand)):
                if hand[i] > 0:
                    for j in range(i + 1, len(hand)):
                        if hand[j] > 0:
                            Relations[hand[i] - 1][hand[j] + 1] += 1
                            Relations[hand[j] - 1][hand[i] + 1] += 1
#       PrintList(Relations,0,3)
        #        PrintList(Hands,3,3)

        with open(path + "\\relations" + str(int(minhands)) + ".txt", "w", encoding="utf8") as relationsfile:
            for relation in Relations:
                for item in relation:
                    relationsfile.write(str(item) + " ")
                relationsfile.write("\n")

        Groups = [[], []]
        for w in range(len(Relations)):
            if IdToCheck == 0 or IdToCheck == Relations[w][0]:
                flag = 0
                if IdToCheck == 0:
                    k = w + 2
                else:
                    k = 2
                while k < len(Relations) + 2:
                    if w != k - 2:
                        if Relations[w][k] == 0:
                            if flag == 0:
                                Groups[0].append([w])
                                flag = 1
                            Groups[0][len(Groups[0]) - 1].append(k - 2)
                    k += 1
        for item in Groups[0]:
            #           if item[0] == 1:
            #               sheet.getCellByPosition(0, 76).Value = item[0]
            #               sheet.getCellByPosition(0, 77).Value = len(item)
            for i in range(1, len(item) - 1):
                for j in range(i + 1, len(item)):
                    #                    if item[0] == 1 and item[i] == 3:
                    #                        sheet.getCellByPosition(i+j, 76).Value = item[i]
                    #                        sheet.getCellByPosition(i+j, 77).Value = item[j]
                    #                        sheet.getCellByPosition(i + j, 78).Value = Relations[item[i]][item[j]+2]
                    if Relations[item[i]][item[j] + 2] == 0:
                        Groups[1].append([item[0], item[i], item[j]])
        GroupsID = Groups
        for n in range(len(Groups)):
            for w in range(len(Groups[n])):
                for k in range(len(Groups[n][w])):
                    GroupsID[n][w][k] = Relations[GroupsID[n][w][k]][0]
        #        PrintList(Groups, 0, 29)
        sheet.getCellByPosition(11, 1).Value = min(len(Relations) + 2, 100)
        PrintList(GroupsID[0], 0, 3)
        PrintList(GroupsID[1], 0, 4 + len(GroupsID[0]))
        sheet.getCellByPosition(1, 1).Value = (len(Groups[0]) + len(Groups[1]) + 5)
        sheet.getCellByPosition(11, 1).Value = min(len(Relations) + 2, 100)
        sheet.getCellByPosition(4, 1).Value = dircounter
        sheet.getCellByPosition(2, 1).Value = hands
        sheet.getCellByPosition(3, 1).Value = len(PlayersByHands)
        sheet.getCellByPosition(14, 5).String = "Tworzenie aliasów:"
        sheet.getCellByPosition(15, 5).Value = time.time() - timer
        timer = time.time()

    sheet.getCellByPosition(14, 6).String = "Całość:"
    sheet.getCellByPosition(15, 6).Value = time.time() - start

#        sheet.getCellByPosition(0, 43).Value = len(Groups[1])
