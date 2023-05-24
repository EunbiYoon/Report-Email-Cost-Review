import pandas as pd
import openpyxl
import datetime
import numpy as np

#This weeknum
this_week="23.05 W4"
next_month1="23.06"
next_month2="23.07"

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

# add the expected value
F_BPAE_Merge[next_month1]=F_BPAE_Merge[this_week]-0.5
F_BPAE_Merge[next_month2]=F_BPAE_Merge[this_week]-1

F_PACE_Merge[next_month1]=F_PACE_Merge[this_week]-0.5
F_PACE_Merge[next_month2]=F_PACE_Merge[this_week]-1

T_BPAE_Merge[next_month1]=T_BPAE_Merge[this_week]-0.5
T_BPAE_Merge[next_month2]=T_BPAE_Merge[this_week]-1

T_PACE_Merge[next_month1]=T_PACE_Merge[this_week]-0.5
T_PACE_Merge[next_month2]=T_PACE_Merge[this_week]-1

D_BPAE_Merge[next_month1]=D_BPAE_Merge[this_week]-0.5
D_BPAE_Merge[next_month2]=D_BPAE_Merge[this_week]-1

D_PACE_Merge[next_month1]=D_PACE_Merge[this_week]-0.5
D_PACE_Merge[next_month2]=D_PACE_Merge[this_week]-1


# change index
F_BPAE_Merge.index=range(1,len(F_BPAE_Merge)+1)
F_PACE_Merge.index=range(1,len(F_PACE_Merge)+1)
T_BPAE_Merge.index=range(1,len(T_BPAE_Merge)+1)
T_PACE_Merge.index=range(1,len(T_PACE_Merge)+1)
D_BPAE_Merge.index=range(1,len(D_BPAE_Merge)+1)
D_PACE_Merge.index=range(1,len(D_PACE_Merge)+1)
############################ Item Table ############################  
#data read
F_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA FL")
D_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA Dryer")
T_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="BPA TL")

F_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="FL_PAC_Item2")
D_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="DR_PAC_Item2")
T_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0526/data.xlsx", sheet_name="TL_PAC_Item")


#data required
F_BPA_I=F_BPA_I.drop([0],axis=0)
F_BPA_I=F_BPA_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
F_BPA_I=F_BPA_I.dropna()
F_BPA_I=F_BPA_I.drop(['Gap'],axis=1)
F_BPA_I=F_BPA_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
F_BPA_I.reset_index(drop=True, inplace=True)

D_BPA_I=D_BPA_I.drop([0],axis=0)
D_BPA_I=D_BPA_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
D_BPA_I=D_BPA_I.dropna()
D_BPA_I=D_BPA_I.drop(['Gap'],axis=1)
D_BPA_I=D_BPA_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
D_BPA_I.reset_index(drop=True, inplace=True)

T_BPA_I=T_BPA_I.drop([0],axis=0)
T_BPA_I=T_BPA_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
T_BPA_I=T_BPA_I.dropna()
T_BPA_I=T_BPA_I.drop(['Gap'],axis=1)
T_BPA_I=T_BPA_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
T_BPA_I.reset_index(drop=True, inplace=True)

F_PAC_I=F_PAC_I.drop([0],axis=0)
F_PAC_I=F_PAC_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
F_PAC_I=F_PAC_I.dropna()
F_PAC_I=F_PAC_I.drop(['Gap'],axis=1)
F_PAC_I=F_PAC_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
F_PAC_I.reset_index(drop=True, inplace=True)

D_PAC_I=D_PAC_I.drop([0],axis=0)
D_PAC_I=D_PAC_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
D_PAC_I=D_PAC_I.dropna()
D_PAC_I=D_PAC_I.drop(['Gap'],axis=1)
D_PAC_I=D_PAC_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
D_PAC_I.reset_index(drop=True, inplace=True)

T_PAC_I=T_PAC_I.drop([0],axis=0)
T_PAC_I=T_PAC_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
T_PAC_I=T_PAC_I.dropna()
T_PAC_I=T_PAC_I.drop(['Gap'],axis=1)
T_PAC_I=T_PAC_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
T_PAC_I.reset_index(drop=True, inplace=True)


# part number same merge
FBI=pd.DataFrame(F_BPA_I.groupby(['Part No','Description']).sum())
FBI.reset_index(inplace=True)
TBI=pd.DataFrame(T_BPA_I.groupby(['Part No','Description']).sum())
TBI.reset_index(inplace=True)
DBI=pd.DataFrame(D_BPA_I.groupby(['Part No','Description']).sum())
DBI.reset_index(inplace=True)
FPI=pd.DataFrame(F_PAC_I.groupby(['Part No','Description']).sum())
FPI.reset_index(inplace=True)
TPI=pd.DataFrame(T_PAC_I.groupby(['Part No','Description']).sum())
TPI.reset_index(inplace=True)
DPI=pd.DataFrame(D_PAC_I.groupby(['Part No','Description']).sum())
DPI.reset_index(inplace=True)


#sort top 5 (increase:2, decrease :3)
FBI=FBI.sort_values(by='VI',ascending=True)
FBI_L=FBI[:3]
FBI=FBI.sort_values(by='VI',ascending=False)
FBI_H=FBI[:2]

TBI=TBI.sort_values(by='VI',ascending=True)
TBI_L=TBI[:3]
TBI=TBI.sort_values(by='VI',ascending=False)
TBI_H=TBI[:2]

DBI=DBI.sort_values(by='VI',ascending=True)
DBI_L=DBI[:3]
DBI=DBI.sort_values(by='VI',ascending=False)
DBI_H=DBI[:2]

FPI=FPI.sort_values(by='VI',ascending=True)
FPI_L=FPI[:3]
FPI=FPI.sort_values(by='VI',ascending=False)
FPI_H=FPI[:2]

TPI=TPI.sort_values(by='VI',ascending=True)
TPI_L=TPI[:3]
TPI=TPI.sort_values(by='VI',ascending=False)
TPI_H=TPI[:2]

DPI=DPI.sort_values(by='VI',ascending=True)
DPI_L=DPI[:3]
DPI=DPI.sort_values(by='VI',ascending=False)
DPI_H=DPI[:2]


#high value need to check it is all increase or not,
FBI_H.reset_index(inplace=True, drop=True)
for i in range(len(FBI_H)):
    condition=FBI_H.at[i,"VI"]
    if condition<0:
        FBI_H=FBI_H.drop([i],axis=0)

TBI_H.reset_index(inplace=True, drop=True)
for i in range(len(FBI_H)):
    condition=TBI_H.at[i,"VI"]
    if condition<0:
        TBI_H=TBI_H.drop([i],axis=0)

DBI_H.reset_index(inplace=True, drop=True)
for i in range(len(DBI_H)):
    condition=DBI_H.at[i,"VI"]
    if condition<0:
        DBI_H=DBI_H.drop([i],axis=0)

FPI_H.reset_index(inplace=True, drop=True)
for i in range(len(FPI_H)):
    condition=FPI_H.at[i,"VI"]
    if condition<0:
        FPI_H=FPI_H.drop([i],axis=0)

TPI_H.reset_index(inplace=True, drop=True)
for i in range(len(TPI_H)):
    condition=TPI_H.at[i,"VI"]
    if condition<0:
        TPI_H=TPI_H.drop([i],axis=0)

DPI_H.reset_index(inplace=True, drop=True)
for i in range(len(DPI_H)):
    condition=DPI_H.at[i,"VI"]
    if condition<0:
        DPI_H=DPI_H.drop([i],axis=0)

#reset H and L
FBI_H.reset_index(inplace=True, drop=True)
TBI_H.reset_index(inplace=True, drop=True)
DBI_H.reset_index(inplace=True, drop=True)
FPI_H.reset_index(inplace=True, drop=True)
TPI_H.reset_index(inplace=True, drop=True)
DPI_H.reset_index(inplace=True, drop=True)

FBI_L.reset_index(inplace=True, drop=True)
TBI_L.reset_index(inplace=True, drop=True)
DBI_L.reset_index(inplace=True, drop=True)
FPI_L.reset_index(inplace=True, drop=True)
TPI_L.reset_index(inplace=True, drop=True)
DPI_L.reset_index(inplace=True, drop=True)

# read previous report and merge
FBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="FL_BPA_Item")
TBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="TL_BPA_Item")
DBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="DR_BPA_Item")
FPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="FL_PAC_Item")
TPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="TL_PAC_Item")
DPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/Cost Review_0519.xlsx", sheet_name="DR_PAC_Item")

FBI_P.index=FBI_P["Unnamed: 0"]
FBI_P=FBI_P.drop(["Unnamed: 0"], axis=1)
FBI_P=FBI_P.drop(["Total"], axis=0)
FBI_P=FBI_P.rename_axis(None)
FBI_P_1=FBI_P[['Increase','VI','Date']].dropna()
FBI_P_2=FBI_P[['Decrease','VI.1','Date.1']].dropna()

TBI_P.index=TBI_P["Unnamed: 0"]
TBI_P=TBI_P.drop(["Unnamed: 0"], axis=1)
TBI_P=TBI_P.drop(["Total"], axis=0)
TBI_P=TBI_P.rename_axis(None)
TBI_P_1=TBI_P[['Increase','VI','Date']].dropna()
TBI_P_2=TBI_P[['Decrease','VI.1','Date.1']].dropna()

DBI_P.index=DBI_P["Unnamed: 0"]
DBI_P=DBI_P.drop(["Unnamed: 0"], axis=1)
DBI_P=DBI_P.drop(["Total"], axis=0)
DBI_P=DBI_P.rename_axis(None)
DBI_P_1=DBI_P[['Increase','VI','Date']].dropna()
DBI_P_2=DBI_P[['Decrease','VI.1','Date.1']].dropna()

FPI_P.index=FPI_P["Unnamed: 0"]
FPI_P=FPI_P.drop(["Unnamed: 0"], axis=1)
FPI_P=FPI_P.drop(["Total"], axis=0)
FPI_P=FPI_P.rename_axis(None)
FPI_P_1=FPI_P[['Increase','VI','Date']].dropna()
FPI_P_2=FPI_P[['Decrease','VI.1','Date.1']].dropna()

TPI_P.index=TPI_P["Unnamed: 0"]
TPI_P=TPI_P.drop(["Unnamed: 0"], axis=1)
TPI_P=TPI_P.drop(["Total"], axis=0)
TPI_P=TPI_P.rename_axis(None)
TPI_P_1=TPI_P[['Increase','VI','Date']].dropna()
TPI_P_2=TPI_P[['Decrease','VI.1','Date.1']].dropna()

DPI_P.index=DPI_P["Unnamed: 0"]
DPI_P=DPI_P.drop(["Unnamed: 0"], axis=1)
DPI_P=DPI_P.drop(["Total"], axis=0)
DPI_P=DPI_P.rename_axis(None)
DPI_P_1=DPI_P[['Increase','VI','Date']].dropna()
DPI_P_2=DPI_P[['Decrease','VI.1','Date.1']].dropna()


# insert to old data
row_1=len(FBI_P_1)
row_2=len(FBI_P_2)
for i in range(len(FBI_H)):
    row_1=row_1+1
    FBI_P_1.at[row_1,"Increase"]=FBI_H.at[i,"Description"]
    FBI_P_1.at[row_1,"VI"]=FBI_H.at[i,"VI"]
    FBI_P_1.at[row_1,"Date"]=this_week
for i in range(len(FBI_L)):
    row_2=row_2+1
    FBI_P_2.at[row_2,"Decrease"]=FBI_L.at[i,"Description"]
    FBI_P_2.at[row_2,"VI.1"]=FBI_L.at[i,"VI"]
    FBI_P_2.at[row_2,"Date.1"]=this_week


row_1=len(TBI_P_1)
row_2=len(TBI_P_2)
for i in range(len(TBI_H)):
    row_1=row_1+1
    TBI_P_1.at[row_1,"Increase"]=TBI_H.at[i,"Description"]
    TBI_P_1.at[row_1,"VI"]=TBI_H.at[i,"VI"]
    TBI_P_1.at[row_1,"Date"]=this_week
for i in range(len(TBI_L)):
    row_2=row_2+1
    TBI_P_2.at[row_2,"Decrease"]=TBI_L.at[i,"Description"]
    TBI_P_2.at[row_2,"VI.1"]=TBI_L.at[i,"VI"]
    TBI_P_2.at[row_2,"Date.1"]=this_week


row_1=len(DBI_P_1)
row_2=len(DBI_P_2)
for i in range(len(DBI_H)):
    row_1=row_1+1
    DBI_P_1.at[row_1,"Increase"]=DBI_H.at[i,"Description"]
    DBI_P_1.at[row_1,"VI"]=DBI_H.at[i,"VI"]
    DBI_P_1.at[row_1,"Date"]=this_week

for i in range(len(DBI_L)):
    row_2=row_2+1
    DBI_P_2.at[row_2,"Decrease"]=DBI_L.at[i,"Description"]
    DBI_P_2.at[row_2,"VI.1"]=DBI_L.at[i,"VI"]
    DBI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(FPI_P_1)
row_2=len(FPI_P_2)
for i in range(len(FPI_H)):
    row_1=row_1+1
    FPI_P_1.at[row_1,"Increase"]=FPI_H.at[i,"Description"]
    FPI_P_1.at[row_1,"VI"]=FPI_H.at[i,"VI"]
    FPI_P_1.at[row_1,"Date"]=this_week
for i in range(len(FPI_L)):
    row_2=row_2+1
    FPI_P_2.at[row_2,"Decrease"]=FPI_L.at[i,"Description"]
    FPI_P_2.at[row_2,"VI.1"]=FPI_L.at[i,"VI"]
    FPI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(TPI_P_1)
row_2=len(TPI_P_2)
for i in range(len(TPI_H)):
    row_1=row_1+1
    TPI_P_1.at[row_1,"Increase"]=TPI_H.at[i,"Description"]
    TPI_P_1.at[row_1,"VI"]=TPI_H.at[i,"VI"]
    TPI_P_1.at[row_1,"Date"]=this_week
for i in range(len(TPI_L)):
    row_2=row_2+1
    TPI_P_2.at[row_2,"Decrease"]=TPI_L.at[i,"Description"]
    TPI_P_2.at[row_2,"VI.1"]=TPI_L.at[i,"VI"]
    TPI_P_2.at[row_2,"Date.1"]=this_week


row_1=len(DPI_P_1)
row_2=len(DPI_P_2)
for i in range(len(DPI_H)):
    row_1=row_1+1
    DPI_P_1.at[row_1,"Increase"]=DPI_H.at[i,"Description"]
    DPI_P_1.at[row_1,"VI"]=DPI_H.at[i,"VI"]
    DPI_P_1.at[row_1,"Date"]=this_week
for i in range(len(DPI_L)):
    row_2=row_2+1
    DPI_P_2.at[row_2,"Decrease"]=DPI_L.at[i,"Description"]
    DPI_P_2.at[row_2,"VI.1"]=DPI_L.at[i,"VI"]
    DPI_P_2.at[row_2,"Date.1"]=this_week


#merge increase and decrease again
FBI_merge = pd.concat([FBI_P_1, FBI_P_2], axis=1)
TBI_merge = pd.concat([TBI_P_1, TBI_P_2], axis=1)
DBI_merge = pd.concat([DBI_P_1, DBI_P_2], axis=1)
FPI_merge = pd.concat([FPI_P_1, FPI_P_2], axis=1)
TPI_merge = pd.concat([TPI_P_1, TPI_P_2], axis=1)
DPI_merge = pd.concat([DPI_P_1, DPI_P_2], axis=1)


#Total value
FBI_sum=FBI_merge.sum()
FBI_merge.at["Total","VI"]=FBI_sum["VI"]
FBI_merge.at["Total","VI.1"]=FBI_sum["VI.1"]

TBI_sum=TBI_merge.sum()
TBI_merge.at["Total","VI"]=TBI_sum["VI"]
TBI_merge.at["Total","VI.1"]=TBI_sum["VI.1"]

DBI_sum=DBI_merge.sum()
DBI_merge.at["Total","VI"]=DBI_sum["VI"]
DBI_merge.at["Total","VI.1"]=DBI_sum["VI.1"]

FPI_sum=FPI_merge.sum()
FPI_merge.at["Total","VI"]=FPI_sum["VI"]
FPI_merge.at["Total","VI.1"]=FPI_sum["VI.1"]

TPI_sum=TPI_merge.sum()
TPI_merge.at["Total","VI"]=TPI_sum["VI"]
TPI_merge.at["Total","VI.1"]=TPI_sum["VI.1"]

DPI_sum=DPI_merge.sum()
DPI_merge.at["Total","VI"]=DPI_sum["VI"]
DPI_merge.at["Total","VI.1"]=DPI_sum["VI.1"]

#round2
FBI_merge=FBI_merge.round(2)
TBI_merge=TBI_merge.round(2)
DBI_merge=DBI_merge.round(2)
FPI_merge=FPI_merge.round(2)
TPI_merge=TPI_merge.round(2)
DPI_merge=DPI_merge.round(2)

# #delete index column
# FBI_merge=pd.DataFrame(FBI_merge)
# print(FBI_merge)
# TBI_merge=TBI_merge.reset_index(range(1,len(TBI_merge)+1))
# DBI_merge=DBI_merge.reset_index(range(1,len(DBI_merge)+1))
# FPI_merge=FPI_merge.reset_index(range(1,len(FPI_merge)+1))
# TPI_merge=TPI_merge.reset_index(range(1,len(TPI_merge)+1))
# DPI_merge=DPI_merge.reset_index(range(1,len(DPI_merge)+1))

# #delete index column
# FBI_merge=FBI_merge.drop(["Index"],axis=1)
# TBI_merge=TBI_merge.drop(["Index"],axis=1)
# DBI_merge=DBI_merge.drop(["Index"],axis=1)
# FPI_merge=FPI_merge.drop(["Index"],axis=1)
# TPI_merge=TPI_merge.drop(["Index"],axis=1)
# DPI_merge=DPI_merge.drop(["Index"],axis=1)

#nan -> blank
FBI_merge = FBI_merge.replace(np.nan, '', regex=True)
TBI_merge = TBI_merge.replace(np.nan, '', regex=True)
DBI_merge = DBI_merge.replace(np.nan, '', regex=True)
FPI_merge = FPI_merge.replace(np.nan, '', regex=True)
TPI_merge = TPI_merge.replace(np.nan, '', regex=True)
DPI_merge = DPI_merge.replace(np.nan, '', regex=True)

#column name change vi.1 and Date.1
FBI_merge=FBI_merge.rename(columns={"VI.1": "VI", "Date.1": "Date"})
TBI_merge=TBI_merge.rename(columns={"VI.1": "VI", "Date.1": "Date"})
DBI_merge=DBI_merge.rename(columns={"VI.1": "VI", "Date.1": "Date"})
FPI_merge=FPI_merge.rename(columns={"VI.1": "VI", "Date.1": "Date"})
TPI_merge=TPI_merge.rename(columns={"VI.1": "VI", "Date.1": "Date"})
DPI_merge=DPI_merge.rename(columns={"VI.1": "VI", "Date.1": "Date"})


############################ Write excel ############################  
file_writer = pd.ExcelWriter("C:/Users/RnD Workstation/Documents/CostReview/0526/Cost Review_0526.xlsx", engine="xlsxwriter")

F_BPAE_Merge.to_excel(file_writer, sheet_name="FL_BPA")
FBI_merge.to_excel(file_writer, sheet_name="FL_BPA_Item")

T_BPAE_Merge.to_excel(file_writer, sheet_name="TL_BPA")
TBI_merge.to_excel(file_writer, sheet_name="TL_BPA_Item")

D_BPAE_Merge.to_excel(file_writer, sheet_name="DR_BPA")
DBI_merge.to_excel(file_writer, sheet_name="DR_BPA_Item")

F_PACE_Merge.to_excel(file_writer, sheet_name="FL_PAC")
FPI_merge.to_excel(file_writer, sheet_name="FL_PAC_Item")

T_PACE_Merge.to_excel(file_writer, sheet_name="TL_PAC")
TPI_merge.to_excel(file_writer, sheet_name="TL_PAC_Item")

D_PACE_Merge.to_excel(file_writer, sheet_name="DR_PAC")
DPI_merge.to_excel(file_writer, sheet_name="DL_PAC_Item")

file_writer.close()

############################ html  ############################  
# change to html -> table & border
F_BPAE_html=F_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')
T_BPAE_html=T_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:sans-serif;"')
D_BPAE_html=D_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:sans-serif;"')
F_PACE_html=F_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:sans-serif;"')
T_PACE_html=T_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:sans-serif;"')
D_PACE_html=D_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:sans-serif;"')
FBI_html=FBI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')
TBI_html=TBI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')
DBI_html=DBI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')
FPI_html=FPI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')
TPI_html=TPI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')
DPI_html=DPI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:sans-serif;"')


# text align center & column color
F_BPAE_html=F_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_BPAE_html=T_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_BPAE_html=D_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
F_PACE_html=F_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_PACE_html=T_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_PACE_html=D_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')

FBI_html=FBI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TBI_html=TBI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DBI_html=DBI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
FPI_html=FPI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TPI_html=TPI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DPI_html=DPI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')


# row color & padding
F_BPAE_html=F_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_BPAE_html=T_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_BPAE_html=D_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
F_PACE_html=F_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_PACE_html=T_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_PACE_html=D_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
FBI_html=FBI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TBI_html=TBI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DBI_html=DBI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
FPI_html=FPI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TPI_html=TPI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DPI_html=DPI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')


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


#this week remark
thisweek_html='<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">'+this_week+'</th>'
thisweek_strong='<th style="text-align: center; background-color:rgb(192, 0, 0); color:white; padding:5px;">'+this_week+'</th>'
F_BPAE_html=F_BPAE_html.replace(thisweek_html,thisweek_strong)
T_BPAE_html=T_BPAE_html.replace(thisweek_html,thisweek_strong)
D_BPAE_html=D_BPAE_html.replace(thisweek_html,thisweek_strong)
F_PACE_html=F_PACE_html.replace(thisweek_html,thisweek_strong)
T_PACE_html=T_PACE_html.replace(thisweek_html,thisweek_strong)
D_PACE_html=D_PACE_html.replace(thisweek_html,thisweek_strong)
