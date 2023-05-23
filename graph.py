from table import F_BPAE_Merge,T_BPAE_Merge,D_BPAE_Merge,F_PACE_Merge,T_PACE_Merge,D_PACE_Merge
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
import random
from random import randrange
import numpy as np

#save file location
save_path="C:/Users/RnD Workstation/Documents/CostReview/0526/"
today_date="_0526"
color_list=['#4472C4','#A5A5A5','#70AD47','#FFC000','#ED7D31','#6600CC','#FF66FF']

########################################## FL - BPA Entity ##########################################
#change point
graph_data=F_BPAE_Merge
graph_title='Front Loader BPA Entity Trend Graph'
file_name="FL_BPA_Entity"+today_date

#data 가공
graph_column=list(graph_data["Model"])
graph_data.reset_index(inplace=True, drop=True)
graph_data=graph_data.drop(["Model","Tool"],axis=1)
graph_data=graph_data.T
graph_data.index=graph_data.index.astype(str)

# graph 생성
fig, ax = plt.subplots()
fig.set_size_inches(15, 4)
for i in range(len(graph_column)):
    grap_list=graph_data[[i]]
    ax.plot(grap_list,linestyle='-',linewidth=1.0,label=graph_column[i],color=color_list[i]) # SVC
# ax.set_title(graph_title,pad=0.2,loc='center',size="13",weight="800", family="sans-serif")
ax.yaxis.tick_right()
ax.legend()
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)


########################################## TL - BPA Entity ##########################################
#change point
graph_data=T_BPAE_Merge
graph_title='Top Loader BPA Entity Trend Graph'
file_name="TL_BPA_Entity"+today_date

#data 가공
graph_column=list(graph_data["Model"])
graph_data.reset_index(inplace=True, drop=True)
graph_data=graph_data.drop(["Model","Tool"],axis=1)
graph_data=graph_data.T
graph_data.index=graph_data.index.astype(str)

# graph 생성
fig, ax = plt.subplots()
fig.set_size_inches(15, 4)
for i in range(len(graph_column)):
    grap_list=graph_data[[i]]
    ax.plot(grap_list,linestyle='-', label=graph_column[i],linewidth=1.0,color=color_list[i]) # SVC
# ax.set_title(graph_title,pad=0.2,loc='center',size="13",weight="800", family="sans-serif")
ax.yaxis.tick_right()
ax.legend()
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)

########################################## DR - BPA Entity ##########################################
#change point
graph_data=D_BPAE_Merge
graph_title='Dryer BPA Entity Trend Graph'
file_name="DR_BPA_Entity"+today_date

#data 가공
graph_column=list(graph_data["Model"])
graph_data.reset_index(inplace=True, drop=True)
graph_data=graph_data.drop(["Model","Tool"],axis=1)
graph_data=graph_data.T
graph_data.index=graph_data.index.astype(str)

# graph 생성
fig, ax = plt.subplots()
fig.set_size_inches(15, 4)
for i in range(len(graph_column)):
    grap_list=graph_data[[i]]
    ax.plot(grap_list,linestyle='-', label=graph_column[i],linewidth=1.0,color=color_list[i]) # SVC
# ax.set_title(graph_title,pad=0.2,loc='center',size="13",weight="800", family="sans-serif")
ax.yaxis.tick_right()
ax.legend()
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)

########################################## FL - PAC Entity ##########################################
#change point
graph_data=F_PACE_Merge
graph_title='Front Loader PAC Entity Trend Graph'
file_name="FL_PAC_Entity"+today_date

#data 가공
graph_column=list(graph_data["Model"])
graph_data.reset_index(inplace=True, drop=True)
graph_data=graph_data.drop(["Model","Tool"],axis=1)
graph_data=graph_data.T
graph_data.index=graph_data.index.astype(str)

# graph 생성
fig, ax = plt.subplots()
fig.set_size_inches(15, 4)
for i in range(len(graph_column)):
    grap_list=graph_data[[i]]
    ax.plot(grap_list,linestyle='-', label=graph_column[i],linewidth=1.0,color=color_list[i]) # SVC
# ax.set_title(graph_title,pad=0.2,loc='center',size="13",weight="800", family="sans-serif")
ax.yaxis.tick_right()
ax.legend()
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)

########################################## TL - PAC Entity ##########################################
#change point
graph_data=T_PACE_Merge
graph_title='Top Loader PAC Entity Trend Graph'
file_name="TL_PAC_Entity"+today_date

#data 가공
graph_column=list(graph_data["Model"])
graph_data.reset_index(inplace=True, drop=True)
graph_data=graph_data.drop(["Model","Tool"],axis=1)
graph_data=graph_data.T
graph_data.index=graph_data.index.astype(str)

# graph 생성
fig, ax = plt.subplots()
fig.set_size_inches(15, 4)
for i in range(len(graph_column)):
    grap_list=graph_data[[i]]
    ax.plot(grap_list,linestyle='-', label=graph_column[i],linewidth=1.0,color=color_list[i]) # SVC
# ax.set_title(graph_title,pad=0.2,loc='center',size="13",weight="800", family="sans-serif")
ax.yaxis.tick_right()
ax.legend()
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)

########################################## DR - PAC Entity ##########################################
#change point
graph_data=D_PACE_Merge
graph_title='Dryer PAC Entity Trend Graph'
file_name="DR_PAC_Entity"+today_date

#data 가공
graph_column=list(graph_data["Model"])
graph_data.reset_index(inplace=True, drop=True)
graph_data=graph_data.drop(["Model","Tool"],axis=1)
graph_data=graph_data.T
graph_data.index=graph_data.index.astype(str)

# graph 생성
fig, ax = plt.subplots()
fig.set_size_inches(15, 4)
for i in range(len(graph_column)):
    grap_list=graph_data[[i]]
    ax.plot(grap_list,linestyle='-', label=graph_column[i],linewidth=1.0,color=color_list[i]) # SVC
# ax.set_title(graph_title,pad=0.2,loc='center',size="13",weight="800", family="sans-serif")
ax.yaxis.tick_right()
ax.legend()
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)

