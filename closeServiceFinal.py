"""
Y.Wang
Date: 23.09.2021
"""
import os
import time
import datetime

def closeService():
    i = 1
    while i == 3:
        print(i)
        i += 1
        time.sleep(1)
    os.system(r'C:\Users\User\Desktop\close.bat')

sched_time = datetime.datetime(2021, 9, 24, 10, 50, 0)
loopflag = 0
while True:
    now = datetime.datetime.now()
    if sched_time<now<(sched_time+datetime.timedelta(seconds=1)):
        loopflag = 1
        time.sleep(1)
    if loopflag == 1:
    # if now == sched_time:
        closeService()
    loopflag = 0