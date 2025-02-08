# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
# import os
# import subprocess
import pandas as pd
import random
import time

# path = os.getcwd() + "/hello.txt"
# subprocess.Popen(["notepad.exe"])

df = pd.read_csv("1000.csv", header = None)
# print(df.head())

# print(df.columns)
# print(df.shape)

# dt = df[0][1]
# print(dt)


# l = ['Mensch', 'Ferse']
# a = df[df[0].isin(l)]
# print(a)

n = df.shape[0]
# print(n)

for _ in range(20):
    r = random.randint(0, n - 1)
    print(df.iloc[r][0])
    #query = input("What does it mean?") #manuel ver
    time.sleep(3) #automatic ver
    b,c = df.iloc[r][1:3]
    print(b,c)
    print("---------------------------------------------------------------")
    time.sleep(2)