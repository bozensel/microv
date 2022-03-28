import pandas as pd
import numpy as np
import xlrd
import openpyxl
import os
from pandas import DataFrame

with open("ABC.csv",'r') as f:
    with open("area1.csv",'w') as f1:
        next(f) 
        for line in f:
            f1.write(line)

with open("DEF.csv",'r') as f:
    with open("area2.csv",'w') as f1:
        next(f)
        for line in f:
            f1.write(line)

with open("JKL.csv",'r') as f:
    with open("area3.csv",'w') as f1:
        next(f) 
        for line in f:
            f1.write(line)

with open("MNO",'r') as f:
    with open("area4.csv",'w') as f1:
        next(f)
        for line in f:
            f1.write(line)


#pre limitation
pd.set_option('display.width', 400)
pd.set_option('display.max_columns', 100)

#Desired columns are taken from CSVs
area1 = pd.read_csv('area1.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
#area2 = pd.read_csv('area2.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
#area3 = pd.read_csv('area3.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
#area4 = pd.read_csv('area4.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
#area5 = pd.read_csv('area5.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
area6 = pd.read_csv('area6.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
area7= pd.read_csv('area7.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
area8 = pd.read_csv('area8.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])
#area9 = pd.read_csv('area9.csv',  usecols=[0,1,2,3,4,5,6,7,9,13,17,26,66,91,106])

#area1=area1.append(area2, ignore_index=True)
#area1=area1.append(area3, ignore_index=True)
#area1=area1.append(area4, ignore_index=True)
#area1=area1.append(area5, ignore_index=True)
area1=area1.append(area6, ignore_index=True)
area1=area1.append(area7, ignore_index=True)
area1=area1.append(area8, ignore_index=True)
#area1=area1.append(area9, ignore_index=True)

wrc = pd.ExcelWriter('area1_0.xlsx')
area1.to_excel(wrc, sheet_name='dfs')
wrc.save()

df0= pd.read_excel('area1_0.xlsx')
df0["Terminal_ID"]=df0["Terminal_ID"].astype(str)
dfs=df0["Terminal_ID"].str.slice(1,2,1)
df0.insert(loc=0, column='TYPE',value=dfs)

wr0 = pd.ExcelWriter('area1_1.xlsx')
df0.to_excel(wr0, sheet_name='dfs')
wr0.save()

dft= pd.read_excel('area1_1.xlsx')

dft["TYPE"]=dft["TYPE"].replace(["0"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["1"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["2"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["3"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["4"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["5"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["6"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["7"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["8"],"MOBIL")
dft["TYPE"]=dft["TYPE"].replace(["9"],"MOBIL")

dft["TYPE"]=dft["TYPE"].replace(["A"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["B"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["C"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["D"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["E"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["F"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["G"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["H"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["I"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["J"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["K"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["L"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["M"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["N"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["O"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["P"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["Q"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["R"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["S"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["T"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["U"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["V"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["W"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["X"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["Y"],"CORPORATE")
dft["TYPE"]=dft["TYPE"].replace(["Z"],"CORPORATE")

#Mobile & Enterprise file separation
dfk=dft.loc[dft["TYPE"]=="CORPORATE"]
dfk.drop(["Unnamed: 0","Unnamed: 0.1"],axis=1,inplace=True)
dfm=dft.loc[dft["TYPE"]=="MOBIL"]
dfm.drop(["Unnamed: 0","Unnamed: 0.1"],axis=1,inplace=True)


wr1 = pd.ExcelWriter('dfm.xlsx')
dfm.to_excel(wr1, sheet_name='1+1 & Manual')
wr1.save()


#Equipment_Protection_Mode (1+1) & Manual
df1=dfm
df1=df1.loc[df1["Protection_Mode_Admin_Status"]=="1+1 Hot"]
df1=df1.loc[df1["Equipment_Protection_Mode"]=="manual"]


#print(df1)

#Output_Power_Admin_Status & RTPC
df2=dfm
df2=df2.loc[df2["Output_Power_Admin_Status"]=="RTPC"]

#print(df2)

#Adaptive_Modulation & 4QPSK
df3=dfm
df3=df3.loc[df3["Modulation"]!="4-QAM"]
df3=df3.loc[df3["Modulation"]!="CQPSK"]
df3=df3.loc[df3["Adaptive_Modulation"]=="Auto"]

#print(df3)

#Adaptive_Modulation & Auto
df4=dfm
df4["Type"].fillna("#NA",inplace=True)
df4["Far_End_Type"].fillna("#NA",inplace=True)
df4=df4[df4['Type'].str.contains("MMU3 A|MMU2 H")]
df4=df4[df4['Far_End_Type'].str.contains("MMU3 A|MMU2 H")]
df4=df4.loc[df4["Adaptive_Modulation"]=="Disabled"]


#ATPC_Selected_Input_Power_Far_RF1 & != -30 ve -40
df5=dfm
#df5["Type"].fillna("#NA",inplace=True)
df5=df5.loc[df5['ATPC_Selected_Input_Power_Far_RF1']!=-30]
df5=df5.loc[df5['ATPC_Selected_Input_Power_Far_RF1']!=-40]
#print(df5)


wr1 = pd.ExcelWriter('area1_mobil_check.xlsx')
df1.to_excel(wr1, sheet_name='1+1 & Manual')
df2.to_excel(wr1, sheet_name='RTPC')
df3.to_excel(wr1, sheet_name='MIN_MOD')
df4.to_excel(wr1, sheet_name='AMR AUTO')
df5.to_excel(wr1, sheet_name='Input Power')
wr1.save()

#Equipment_Protection_Mode (1+1) & Manual
df6=dfk
df6=df6.loc[df6["Protection_Mode_Admin_Status"]=="1+1 Hot"]
df6=df6.loc[df6["Equipment_Protection_Mode"]=="manual"]
#print(df1)

#Output_Power_Admin_Status & RTPC
df7=dfk
df7=df7.loc[df7["Output_Power_Admin_Status"]=="RTPC"]
#print(df2)

#ATPC_Selected_Input_Power_Far_RF1 & != -30 ve -40
df8=dfk
#df5["Type"].fillna("#NA",inplace=True)
df8=df8.loc[df8['ATPC_Selected_Input_Power_Far_RF1']!=-30]
df8=df8.loc[df8['ATPC_Selected_Input_Power_Far_RF1']!=-40]
#print(df5)

wr2 = pd.ExcelWriter('area1_CORPORATE_check.xlsx')
df6.to_excel(wr2, sheet_name='1+1 & Manual')
df7.to_excel(wr2, sheet_name='RTPC')
df8.to_excel(wr2, sheet_name='Input Power')
wr2.save()

os.remove("area1_0.xlsx")
os.remove("area1_1.xlsx")
os.remove("area2.xlsx")
