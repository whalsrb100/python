import datetime
def getLastSequence():
    dn=0
    w=0
    today = datetime.datetime.today()
    while w < 6:
        tday = today - datetime.timedelta(days=dn)
        a=int(str(tday.year)[0:2])
        b=int(str(tday.year)[2:5])
        c=int(str(tday.month))
        if c <= 2:
            b = b - 1
            c = c + 12
        d=int(str(tday.day))
        w=int((((21*a/4)+(5*b/4)+(26*(c+1)/10)+d-1)%7))
        dn = dn + 1
    wSeq = 0
    sdate = datetime.date(2002,12,7)
    while sdate.year != tday.year or sdate.month != tday.month or sdate.day != tday.day:
        weak = datetime.timedelta(weeks=wSeq)
        sdate = datetime.date(2002,12,7) + weak
        if today.year == sdate.year and today.month == sdate.month and today.day == sdate.day and today.hour < 20 and today.minute < 50: break
        wSeq = wSeq + 1
    return wSeq