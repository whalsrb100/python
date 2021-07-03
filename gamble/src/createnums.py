from src.score import getValidHistoryDict
def createNums(info,indexPool):
    validHistoryDict = getValidHistoryDict(info)
    scorePerPoint = []
    future = []
    future_tmp = []
    scorepoint = []
    scorepoint_ck = []
    for i in range(0,6):
        future_tmp.append([])
        scorepoint.append(0)
        scorepoint_ck.append(0)
        sumScores = 0
        M = validHistoryDict[i][list(validHistoryDict[i].keys())[0]]
        scorePerPoint.append(int(len(info)/M))
        for j in range(0,len(validHistoryDict[i])):
            scorepoint[i] = int(validHistoryDict[i][list(validHistoryDict[i].keys())[j]] * scorePerPoint[i])
            if scorepoint[i] <= 80 and scorepoint[i] >= 20:
                future_tmp[i].append(list(validHistoryDict[i].keys())[j])
            if len(future_tmp[i]) == 0:
                future_tmp[i].append(list(validHistoryDict[i].keys())[0])

        scorepoint[i] = validHistoryDict[i][future_tmp[i][(indexPool)%len(future_tmp[i])]] * scorePerPoint[i]
    for i in range(0,6):
        for j in range(0,len(info)):
            scorepoint_ck[i] += validHistoryDict[i][info[j][i]] * scorePerPoint[i]
        scorepoint_ck[i] = int(scorepoint_ck[i] /len(info))

    #fSum = rSum = 0
    #for i in range(0,6):
    #    fSum += scorepoint[i]
    #    rSum += scorepoint_ck[i]
    #cha = fSum - rSum
    #if cha < 0: cha *= -1
    #h = 0
    #passNum = 10
    #while cha > passNum:
    #    h += 1
    #    if h > 50000:
    #        passNum += 10
    #        h = 0
    #    fSum=0
    #    for i in range(0,6):
    #        scorepoint[i] = validHistoryDict[i][future_tmp[i][(indexPool+h)%len(future_tmp[i])]] * scorePerPoint[i]
    #        fSum += scorepoint[i]
    #    cha = fSum - rSum
    #    if cha < 0: cha *= -1
    #    #print("h:{}, cha:{}".format(h,cha ) )
    h=0
    for i in range(0,6):
        future.append(future_tmp[i][(indexPool+h)%len(future_tmp[i])])
    for i in range(0,6):
        for j in range(0,6):
            h=0
            if i == j: continue
            while future[i] == future[j]:
                h += 1
                future[j] = future_tmp[j][(indexPool+h)%len(future_tmp[j])]
        #print("r:{}\tf:{}".format(scorepoint_ck[i],scorepoint[i]))


        #cha = scorepoint[i] - scorepoint_ck[i]
        #if cha < 0: cha *=-1
        #h = 1
        #while cha > 30:
        #    h += 1
        #    scorepoint[i] = validHistoryDict[i][future_tmp[i][(indexPool+h)%len(future_tmp[i])]] * scorePerPoint[i]
        #    scorepoint_ck[i] = validHistoryDict[i][info[-1][i]] * scorePerPoint[i]
        #    cha = scorepoint[i] - scorepoint_ck[i]
        #    if cha < 0: cha *=-1
        #future.append(future_tmp[i][(indexPool)%len(future_tmp[i])])


        #print(  "scorepoint:{}\tscorepoint_ck:{}".format( scorepoint[i], scorepoint_ck[i] )  )

        #scorePerPoint.clear()
        #for i in range(0,6):
        #    future_tmp[i].clear()
        #    validHistoryDict[i].clear()
        #validHistoryDict.clear()
        #future_tmp.clear()
        #scorepoint.clear()
        #scorepoint_ck.clear()

    return future
