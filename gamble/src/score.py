# get score function.........
def getValidHistoryDict(info):
    validHistoryDict = []
    for i in range(0,6):
        validHistoryDict.append({})
        for j in range(0, len(info)):
            try: validHistoryDict[i][info[j][i]] += 1
            except KeyError: validHistoryDict[i][info[j][i]] = 1
        validHistoryDict[i] = dict(sorted(validHistoryDict[i].items(), key=(lambda x: x[1]), reverse=True))
    return validHistoryDict
    # validHistoryDict[0-5]{"num":"횟수"}