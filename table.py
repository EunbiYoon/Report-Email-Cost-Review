import pandas as pd
import openpyxl

#read original data
original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/BPAE_0512.xlsx", sheet_name="FL")
original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/PACE_0512.xlsx", sheet_name="FL")

#read new data
bpa_entity=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/BPA_Entity.xlsx")
pac_entity=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0519/PAC_Entity.xlsx")


#remain required  new data
BPAE=bpa_entity[["Model.Suffix","Confirmed\nRMC\nWith Rate\n(USD)"]]
BPAE=BPAE.drop([0])
BPAE.reset_index(inplace=True, drop=True)
BPAE["Confirmed\nRMC\nWith Rate\n(USD)"]=BPAE["Confirmed\nRMC\nWith Rate\n(USD)"].round(1)

pac_entity.columns=pac_entity.iloc[15]
pac_entity=pac_entity.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16],axis=0)
pac_entity.reset_index(drop=True, inplace=True)
PACE=pac_entity[["Model.Suffix","Confirmed\nRMC\nWith Rate\n(USD)"]]
print(PACE)
PACE["Confirmed\nRMC\nWith Rate\n(USD)"]=PACE["Confirmed\nRMC\nWith Rate\n(USD)"].astype(float).round(1)

#Matching column with original data
BPAE=BPAE.rename(columns={"Model.Suffix":"Tool"})
PACE=PACE.rename(columns={"Model.Suffix":"Tool"})


#merge two file with Model Suffix
BPAE_Merge=pd.merge(original_BPAE,BPAE,how="inner",on="Tool")
PACE_Merge=pd.merge(original_PACE,PACE,how="inner",on="Tool")

print(BPAE_Merge)
print(PACE_Merge)
