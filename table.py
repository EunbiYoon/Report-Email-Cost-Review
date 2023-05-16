import pandas as pd
import openpyxl
import datetime

#This weeknum
this_week="May 3rd Week"

############################ Front Loader ############################  
#read original data
F_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/BPAE_0512.xlsx", sheet_name="FL")
F_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/PACE_0512.xlsx", sheet_name="FL")
T_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/BPAE_0512.xlsx", sheet_name="TL")
T_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/PACE_0512.xlsx", sheet_name="TL")
D_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/BPAE_0512.xlsx", sheet_name="DR")
D_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0512/PACE_0512.xlsx", sheet_name="DR")

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
PACE["Confirmed\nRMC\nWith Rate\n(USD)"]=PACE["Confirmed\nRMC\nWith Rate\n(USD)"].astype(float).round(1)


#Matching column with original data
BPAE=BPAE.rename(columns={"Model.Suffix":"Tool","Confirmed\nRMC\nWith Rate\n(USD)":this_week})
PACE=PACE.rename(columns={"Model.Suffix":"Tool","Confirmed\nRMC\nWith Rate\n(USD)":this_week})


#merge two file with Model Suffix
F_BPAE_Merge=pd.merge(F_original_BPAE,BPAE,how="inner",on="Tool")
F_PACE_Merge=pd.merge(F_original_PACE,PACE,how="inner",on="Tool")

T_BPAE_Merge=pd.merge(T_original_BPAE,BPAE,how="inner",on="Tool")
T_PACE_Merge=pd.merge(T_original_PACE,PACE,how="inner",on="Tool")

D_BPAE_Merge=pd.merge(D_original_BPAE,BPAE,how="inner",on="Tool")
D_PACE_Merge=pd.merge(D_original_PACE,PACE,how="inner",on="Tool")


#set index 
F_BPAE_Merge.index=F_BPAE_Merge["Tool"]
F_PACE_Merge.index=F_PACE_Merge["Tool"]

T_BPAE_Merge.index=T_BPAE_Merge["Tool"]
T_PACE_Merge.index=T_PACE_Merge["Tool"]

D_BPAE_Merge.index=D_BPAE_Merge["Tool"]
D_PACE_Merge.index=D_PACE_Merge["Tool"]


#write excel
bpae_writer = pd.ExcelWriter("C:/Users/RnD Workstation/Documents/CostReview/0519/BPAE_0519.xlsx", engine="xlsxwriter")
pace_writer = pd.ExcelWriter("C:/Users/RnD Workstation/Documents/CostReview/0519/PACE_0519.xlsx", engine="xlsxwriter")

F_BPAE_Merge.to_excel(bpae_writer, sheet_name="FL")
T_BPAE_Merge.to_excel(bpae_writer, sheet_name="TL")
D_BPAE_Merge.to_excel(bpae_writer, sheet_name="DR")
bpae_writer.close()

F_PACE_Merge.to_excel(pace_writer, sheet_name="FL")
T_PACE_Merge.to_excel(pace_writer, sheet_name="TL")
D_PACE_Merge.to_excel(pace_writer, sheet_name="DR")
pace_writer.close()


# change to html
F_BPAE_html=F_BPAE_Merge.to_html()
T_BPAE_html=T_BPAE_Merge.to_html()
D_BPAE_html=D_BPAE_Merge.to_html()

F_PACE_html=F_PACE_Merge.to_html()
T_PACE_html=T_PACE_Merge.to_html()
D_PACE_html=D_PACE_Merge.to_html()
