if __name__ == '__main__':
    from getinfo import GetInfo
    from check import Check
    from gethistory import GetHistory

    ginfo = GetInfo(0) # 0: get from local file, 1: get from api
    gck = Check(ginfo.getInfo())
    print('####################### main #######################')
    
    #####################################
    # Set Print boolearn
    #####################################
    bIsPrintCheck=False
    bIsPrintHistory=False
    #####################################

    #####################################
    # print Check class
    #####################################
    if bIsPrintCheck:
        for i in range(len(ginfo.getInfo())):
            print("( {:>4} ) ".format( int(i) ),end='')
            for j in range(6):
                print( "{:>3}, ".format( int(gck.info[i][j]) ), end='' )
            print( "\tbonus:{:>3}".format(int(gck.info[i][6])) )
    #####################################
    
    #####################################
    # print GetHistory class
    #####################################
    if bIsPrintHistory:
        ghist = GetHistory(ginfo.info)
        for i in range(len(ghist.info)):
            print(ghist.info[i])
    #####################################
    print('####################### main #######################')
    del ginfo
    del gck