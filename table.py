import pandas as pd
import openpyxl
import datetime

#This weeknum
this_week="23.05 W4"

############################ Trend Table ############################  
#read original data
F_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="FL_BPA")
F_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="FL_PAC")
T_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="TL_BPA")
T_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="TL_PAC")
D_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="DR_BPA")
D_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="DR_PAC")

#read new data
bpa_entity=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA Entity")
pac_entity=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="PAC Entity")

#remain required  new data
BPAE=bpa_entity[["Model.Suffix","Net RMC (USD)"]]
# BPAE=BPAE.drop([0])
BPAE.reset_index(inplace=True, drop=True)
for i in range(len(BPAE)):
    a=BPAE.at[i,"Net RMC (USD)"]
    BPAE.at[i,"Net RMC (USD)"]=round(a,1)

# pac_entity.columns=pac_entity.iloc[15]
# pac_entity=pac_entity.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16],axis=0)
# pac_entity.reset_index(drop=True, inplace=True)
# print(pac_entity)
# PACE=pac_entity[["Model.Suffix","Net RMC (USD)"]]
# PACE["Net RMC (USD)"]=PACE["Net RMC (USD)"].astype(float).round(1)


PACE=pac_entity[["Model.Suffix","Net RMC (USD)"]]
# PACE=PACE.drop([0])
PACE.reset_index(inplace=True, drop=True)
PACE["Net RMC (USD)"].round(1)
for i in range(len(PACE)):
    a=PACE.at[i,"Net RMC (USD)"]
    PACE.at[i,"Net RMC (USD)"]=round(a,1)


#Matching column with original data
BPAE=BPAE.rename(columns={"Model.Suffix":"Tool","Net RMC (USD)":this_week})
PACE=PACE.rename(columns={"Model.Suffix":"Tool","Net RMC (USD)":this_week})

#merge two file with Model Suffix
F_BPAE_Merge=pd.merge(F_original_BPAE,BPAE,how="inner",on="Tool")
F_PACE_Merge=pd.merge(F_original_PACE,PACE,how="inner",on="Tool")
T_BPAE_Merge=pd.merge(T_original_BPAE,BPAE,how="inner",on="Tool")
T_PACE_Merge=pd.merge(T_original_PACE,PACE,how="inner",on="Tool")
D_BPAE_Merge=pd.merge(D_original_BPAE,BPAE,how="inner",on="Tool")
D_PACE_Merge=pd.merge(D_original_PACE,PACE,how="inner",on="Tool")

#drop unamed:0
F_BPAE_Merge=F_BPAE_Merge.drop(['Unnamed: 0'],axis=1)
F_PACE_Merge=F_PACE_Merge.drop(['Unnamed: 0'],axis=1)
T_BPAE_Merge=T_BPAE_Merge.drop(['Unnamed: 0'],axis=1)
T_PACE_Merge=T_PACE_Merge.drop(['Unnamed: 0'],axis=1)
D_BPAE_Merge=D_BPAE_Merge.drop(['Unnamed: 0'],axis=1)
D_PACE_Merge=D_PACE_Merge.drop(['Unnamed: 0'],axis=1)

# change index
F_BPAE_Merge.index=range(1,len(F_BPAE_Merge)+1)
F_PACE_Merge.index=range(1,len(F_PACE_Merge)+1)
T_BPAE_Merge.index=range(1,len(T_BPAE_Merge)+1)
T_PACE_Merge.index=range(1,len(T_PACE_Merge)+1)
D_BPAE_Merge.index=range(1,len(D_BPAE_Merge)+1)
D_PACE_Merge.index=range(1,len(D_PACE_Merge)+1)

############################ Item Table ############################  
#data read
F_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="FL_PAC_Item")
F_PAC_I2=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="FL_PAC_Item2")
D_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="DR_PAC_Item")
D_PAC_I2=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="DR_PAC_Item2")
T_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="TL_PAC_Item")
T_PAC_I2=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="TL_PAC_Item2")
F_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA FL")
D_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA Dryer")
T_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA TL")

#data required
F_BPA_I=F_BPA_I.drop([0],axis=0)
F_BPA_I=F_BPA_I[['[Left] Model/PartNo.4','Gap.1']]
F_BPA_I=F_BPA_I[2:].reset_index()

D_BPA_I=D_BPA_I.drop([0],axis=0)
D_BPA_I=D_BPA_I[['[Left] Model/PartNo.4','Gap.1']]
D_BPA_I=D_BPA_I[2:].reset_index()

T_BPA_I=T_BPA_I.drop([0],axis=0)
T_BPA_I=T_BPA_I[['[Left] Model/PartNo.4','Gap.1']]
T_BPA_I=T_BPA_I[2:].reset_index()


F_PAC_I=F_PAC_I.drop([0],axis=0)
F_PAC_I=F_PAC_I[['[Left] Model/PartNo.4','Gap.1']]
F_PAC_I=F_PAC_I[2:].reset_index()

F_PAC_I2=F_PAC_I2.drop([0],axis=0)
F_PAC_I2=F_PAC_I2[['[Left] Model/PartNo.4','Gap.1']]
F_PAC_I2=F_PAC_I2[2:].reset_index()

D_PAC_I=D_PAC_I.drop([0],axis=0)
D_PAC_I=D_PAC_I[['[Left] Model/PartNo.4','Gap.1']]
D_PAC_I=D_PAC_I[2:].reset_index()

D_PAC_I2=D_PAC_I2.drop([0],axis=0)
D_PAC_I2=D_PAC_I2[['[Left] Model/PartNo.4','Gap.1']]
D_PAC_I2=D_PAC_I2[2:].reset_index()

T_PAC_I=T_PAC_I.drop([0],axis=0)
T_PAC_I=T_PAC_I[['[Left] Model/PartNo.4','Gap.1']]
T_PAC_I=T_PAC_I[2:].reset_index()

T_PAC_I2=T_PAC_I2.drop([0],axis=0)
T_PAC_I2=T_PAC_I2[['[Left] Model/PartNo.4','Gap.1']]
T_PAC_I2=T_PAC_I2[2:].reset_index()


# 0.5 over sort
FBI=pd.DataFrame()
DBI=pd.DataFrame()
TBI=pd.DataFrame()
FPI=pd.DataFrame()
DPI=pd.DataFrame()
TPI=pd.DataFrame()

for i in range(len(F_BPA_I)):
    condition=abs(F_BPA_I.at[i,'Gap.1'])
    if condition>=0.5:
        FBI.at[i,"Desc"]=F_BPA_I.at[i,'[Left] Model/PartNo.4']
        FBI.at[i,"VI"]=F_BPA_I.at[i,'Gap.1']

for i in range(len(D_BPA_I)):
    condition=abs(D_BPA_I.at[i,'Gap.1'])
    if condition>=0.5:
        DBI.at[i,"Desc"]=D_BPA_I.at[i,'[Left] Model/PartNo.4']
        DBI.at[i,"VI"]=D_BPA_I.at[i,'Gap.1']

for i in range(len(T_BPA_I)):
    condition=abs(T_BPA_I.at[i,'Gap.1'])
    if condition>=0.5:
        TBI.at[i,"Desc"]=T_BPA_I.at[i,'[Left] Model/PartNo.4']
        TBI.at[i,"VI"]=T_BPA_I.at[i,'Gap.1']


# for i in range(len(F_PAC_I)):
#     condition=abs(F_PAC_I.at[i,'Gap.1'])
#     if condition>=0.5:
#         FPI.at[i,"Desc"]=F_PAC_I.at[i,'[Left] Model/PartNo.4']
#         FPI.at[i,"VI"]=F_PAC_I.at[i,'Gap.1']
# FPI_acc=len(FPI)-1
for i in range(len(F_PAC_I2)):
    condition=abs(F_PAC_I2.at[i,'Gap.1'])
    if condition>=0.5:
        FPI.at[i,"Desc"]=F_PAC_I.at[i,'[Left] Model/PartNo.4']
        FPI.at[i,"VI"]=F_PAC_I.at[i,'Gap.1']

# for i in range(len(D_PAC_I)):
#     condition=abs(D_PAC_I.at[i,'Gap.1'])
#     if condition>=0.5:
#         DPI.at[i,"Desc"]=D_PAC_I.at[i,'[Left] Model/PartNo.4']
#         DPI.at[i,"VI"]=D_PAC_I.at[i,'Gap.1']
# DPI_acc=len(DPI)-1
for i in range(len(D_PAC_I2)):
    condition=abs(D_PAC_I2.at[i,'Gap.1'])
    if condition>=0.5:
        DPI.at[i,"Desc"]=D_PAC_I2.at[i,'[Left] Model/PartNo.4']
        DPI.at[i,"VI"]=D_PAC_I2.at[i,'Gap.1']

for i in range(len(T_PAC_I)):
    condition=abs(T_PAC_I.at[i,'Gap.1'])
    if condition>=0.5:
        TPI.at[i,"Desc"]=T_PAC_I.at[i,'[Left] Model/PartNo.4']
        TPI.at[i,"VI"]=T_PAC_I.at[i,'Gap.1']
# TPI_acc=len(TPI)-1
# for i in range(len(T_PAC_I2)):
#     condition=abs(T_PAC_I2.at[i,'Gap.1'])
#     if condition>=0.5:
#         TPI.at[i,"Desc"]=T_PAC_I2.at[i,'[Left] Model/PartNo.4']
#         TPI.at[i,"VI"]=T_PAC_I2.at[i,'Gap.1']

# descend sort
FBI.sort_values(by="VI", ascending=True)
FBI=FBI.reset_index()
FBI=FBI[:3]
TBI.sort_values(by="VI", ascending=True)
TBI=TBI.reset_index()
TBI=TBI[:3]
DBI.sort_values(by="VI", ascending=True)
DBI=DBI.reset_index()
DBI=DBI[:3]
FPI.sort_values(by="VI", ascending=True)
FPI=FPI.reset_index()
FPI=FPI[:3]
TPI.sort_values(by="VI", ascending=True)
TPI=TPI.reset_index()
TPI=TPI[:3]
DPI.sort_values(by="VI", ascending=True)
DPI=DPI.reset_index()
DPI=DPI[:3]

# read previous report and merge
FBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="FL_BPA_Item")
TBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="TL_BPA_Item")
DBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="DR_BPA_Item")
FPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="FL_PAC_Item")
TPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="TL_PAC_Item")
DPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="DR_PAC_Item")

FBI_P.index=FBI_P["Index"]
FBI_P=FBI_P.drop(["Index"], axis=1)
FBI_P=FBI_P.drop(["Total"], axis=0)
FBI_P_1=FBI_P[['Increase','VI','Month']].dropna()
FBI_P_2=FBI_P[['Decrease','VI.1','Month.1']].dropna()

TBI_P.index=TBI_P["Index"]
TBI_P=TBI_P.drop(["Index"], axis=1)
TBI_P=TBI_P.drop(["Total"], axis=0)
TBI_P_1=TBI_P[['Increase','VI','Month']].dropna()
TBI_P_2=TBI_P[['Decrease','VI.1','Month.1']].dropna()

DBI_P.index=DBI_P["Index"]
DBI_P=DBI_P.drop(["Index"], axis=1)
DBI_P=DBI_P.drop(["Total"], axis=0)
DBI_P_1=DBI_P[['Increase','VI','Month']].dropna()
DBI_P_2=DBI_P[['Decrease','VI.1','Month.1']].dropna()

FPI_P.index=FPI_P["Index"]
FPI_P=FPI_P.drop(["Index"], axis=1)
FPI_P=FPI_P.drop(["Total"], axis=0)
FPI_P_1=FPI_P[['Increase','VI','Month']].dropna()
FPI_P_2=FPI_P[['Decrease','VI.1','Month.1']].dropna()

TPI_P.index=TPI_P["Index"]
TPI_P=TPI_P.drop(["Index"], axis=1)
TPI_P=TPI_P.drop(["Total"], axis=0)
TPI_P_1=TPI_P[['Increase','VI','Month']].dropna()
TPI_P_2=TPI_P[['Decrease','VI.1','Month.1']].dropna()

DPI_P.index=DPI_P["Index"]
DPI_P=DPI_P.drop(["Index"], axis=1)
DPI_P=DPI_P.drop(["Total"], axis=0)
DPI_P_1=DPI_P[['Increase','VI','Month']].dropna()
DPI_P_2=DPI_P[['Decrease','VI.1','Month.1']].dropna()


# insert to old data
row_1=len(FBI_P_1)
row_2=len(FBI_P_2)
for i in range(len(FBI)):
    condition=FBI.at[i,"VI"]
    if condition>0:
        FBI_P_1.at[row_1,"Increase"]=FBI.at[i,"Desc"]
        FBI_P_1.at[row_1,"VI"]=FBI.at[i,"VI"]
        FBI_P_1.at[row_1,"Month"]=this_week
        row_1=row_1+1
    else:
        FBI_P_2.at[row_2,"Decrease"]=FBI.at[i,"Desc"]
        FBI_P_2.at[row_2,"VI.1"]=FBI.at[i,"VI"]
        FBI_P_2.at[row_2,"Month.1"]=this_week
        row_2=row_2+1

row_1=len(TBI_P_1)
row_2=len(TBI_P_2)
for i in range(len(TBI)):
    condition=TBI.at[i,"VI"]
    if condition>0:
        TBI_P_1.at[row_1,"Increase"]=TBI.at[i,"Desc"]
        TBI_P_1.at[row_1,"VI"]=TBI.at[i,"VI"]
        TBI_P_1.at[row_1,"Month"]=this_week
        row_1=row_1+1
    else:
        row=len(TBI_P_2)+i
        TBI_P_2.at[row_2,"Decrease"]=TBI.at[i,"Desc"]
        TBI_P_2.at[row_2,"VI.1"]=TBI.at[i,"VI"]
        TBI_P_2.at[row_2,"Month.1"]=this_week
        row_2=row_2+1

row_1=len(DBI_P_1)
row_2=len(DBI_P_2)     
for i in range(len(DBI)):
    condition=DBI.at[i,"VI"]
    if condition>0:
        DBI_P_1.at[row_1,"Increase"]=DBI.at[i,"Desc"]
        DBI_P_1.at[row_1,"VI"]=DBI.at[i,"VI"]
        DBI_P_1.at[row_1,"Month"]=this_week
        row_1=row_1+1
    else:
        row=len(DBI_P_2)+i
        DBI_P_2.at[row_2,"Decrease"]=DBI.at[i,"Desc"]
        DBI_P_2.at[row_2,"VI.1"]=DBI.at[i,"VI"]
        DBI_P_2.at[row_2,"Month.1"]=this_week
        row_2=row_2+1

row_1=len(FPI_P_1)
row_2=len(FPI_P_2) 
for i in range(len(FPI)):
    condition=FPI.at[i,"VI"]
    if condition>0:
        row=len(FPI_P_1)+i
        FPI_P_1.at[row_1,"Increase"]=FPI.at[i,"Desc"]
        FPI_P_1.at[row_1,"VI"]=FPI.at[i,"VI"]
        FPI_P_1.at[row_1,"Month"]=this_week
        row_1=row_1+1
    else:
        FPI_P_2.at[row_2,"Decrease"]=FPI.at[i,"Desc"]
        FPI_P_2.at[row_2,"VI.1"]=FPI.at[i,"VI"]
        FPI_P_2.at[row_2,"Month.1"]=this_week
        row_2=row_2+1

row_1=len(TPI_P_1)
row_2=len(TPI_P_2) 
for i in range(len(TPI)):
    condition=TPI.at[i,"VI"]
    if condition>0:
        TPI_P_1.at[row_1,"Increase"]=TPI.at[i,"Desc"]
        TPI_P_1.at[row_1,"VI"]=TPI.at[i,"VI"]
        TPI_P_1.at[row_1,"Month"]=this_week
        row_1=row_1+1
    else:
        row=len(TPI_P_2)+i
        TPI_P_2.at[row_2,"Decrease"]=TPI.at[i,"Desc"]
        TPI_P_2.at[row_2,"VI.1"]=TPI.at[i,"VI"]
        TPI_P_2.at[row_2,"Month.1"]=this_week
        row_2=row_2+1

row_1=len(DPI_P_1)
row_2=len(DPI_P_2)
for i in range(len(DPI)):
    condition=DPI.at[i,"VI"]
    if condition>0:
        DPI_P_1.at[row_1,"Increase"]=DPI.at[i,"Desc"]
        DPI_P_1.at[row_1,"VI"]=DPI.at[i,"VI"]
        DPI_P_1.at[row_1,"Month"]=this_week
        row_1=row_1+1
    else:
        DPI_P_2.at[row_2,"Decrease"]=DPI.at[i,"Desc"]
        DPI_P_2.at[row_2,"VI.1"]=DPI.at[i,"VI"]
        DPI_P_2.at[row_2,"Month.1"]=this_week
        row_2=row_2+1

#merge increase and decrease again
FBI_merge = pd.concat([FBI_P_1, FBI_P_2], axis=1)
TBI_merge = pd.concat([TBI_P_1, TBI_P_2], axis=1)
DBI_merge = pd.concat([DBI_P_1, DBI_P_2], axis=1)
FPI_merge = pd.concat([FPI_P_1, FPI_P_2], axis=1)
TPI_merge = pd.concat([TPI_P_1, TPI_P_2], axis=1)
DPI_merge = pd.concat([DPI_P_1, DPI_P_2], axis=1)

print(FBI_merge)
print(TBI_merge)
print(DBI_merge)
print(FPI_merge)
print(TPI_merge)
print(DBI_merge)
############################ Write excel ############################  
file_writer = pd.ExcelWriter("C:/Users/RnD Workstation/Documents/CostReview/0526/Cost Review_0526.xlsx", engine="xlsxwriter")
F_BPAE_Merge.to_excel(file_writer, sheet_name="FL_BPA")
T_BPAE_Merge.to_excel(file_writer, sheet_name="TL_BPA")
D_BPAE_Merge.to_excel(file_writer, sheet_name="DR_BPA")
F_PACE_Merge.to_excel(file_writer, sheet_name="FL_PAC")
T_PACE_Merge.to_excel(file_writer, sheet_name="TL_PAC")
D_PACE_Merge.to_excel(file_writer, sheet_name="DR_PAC")
file_writer.close()

############################ html  ############################  
# change to html -> table & border
F_BPAE_html=F_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); border-collapse:collapse; text-align:center;font-family:sans-serif;"')
T_BPAE_html=T_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); border-collapse:collapse; text-align:center;font-family:sans-serif;"')
D_BPAE_html=D_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); border-collapse:collapse; text-align:center;font-family:sans-serif;"')
F_PACE_html=F_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); border-collapse:collapse; text-align:center;font-family:sans-serif;"')
T_PACE_html=T_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); border-collapse:collapse; text-align:center;font-family:sans-serif;"')
D_PACE_html=D_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); border-collapse:collapse; text-align:center;font-family:sans-serif;"')

# text align center & column color
F_BPAE_html=F_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_BPAE_html=T_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_BPAE_html=D_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
F_PACE_html=F_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_PACE_html=T_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_PACE_html=D_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')

# row color & padding
F_BPAE_html=F_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_BPAE_html=T_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_BPAE_html=D_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
F_PACE_html=F_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_PACE_html=T_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_PACE_html=D_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')

# model, tool
F_BPAE_html=F_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
F_BPAE_html=F_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
T_BPAE_html=T_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
T_BPAE_html=T_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
D_BPAE_html=D_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
D_BPAE_html=D_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')

F_PACE_html=F_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
F_PACE_html=F_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
T_PACE_html=T_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
T_PACE_html=T_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
D_PACE_html=D_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
D_PACE_html=D_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
