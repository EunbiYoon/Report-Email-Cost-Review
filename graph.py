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

#value display
graph_data=graph_data.reset_index()
for i in range(len(graph_data.index)):
    ax.annotate(graph_data.at[i,0],xy=(graph_data.at[i,"index"],graph_data.at[i,0]), va='bottom', color='#4472C4',fontsize=9)
    ax.annotate(graph_data.at[i,1],xy=(graph_data.at[i,"index"],graph_data.at[i,1]), va='bottom', color='#A5A5A5',fontsize=9)
    ax.annotate(graph_data.at[i,2],xy=(graph_data.at[i,"index"],graph_data.at[i,2]), va='bottom', color='#70AD47',fontsize=9)
    ax.annotate(graph_data.at[i,3],xy=(graph_data.at[i,"index"],graph_data.at[i,3]), va='top', color='#FFC000',fontsize=9)
    ax.annotate(graph_data.at[i,4],xy=(graph_data.at[i,"index"],graph_data.at[i,4]), va='bottom', color='#ED7D31',fontsize=9)

ax.yaxis.tick_right()
ax.legend(loc='upper center', bbox_to_anchor=(0.5,1.15), ncol=5)
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

#value display
graph_data=graph_data.reset_index()
for i in range(len(graph_data.index)):
    ax.annotate(graph_data.at[i,0],xy=(graph_data.at[i,"index"],graph_data.at[i,0]), va='bottom', color='#4472C4',fontsize=9)
    ax.annotate(graph_data.at[i,1],xy=(graph_data.at[i,"index"],graph_data.at[i,1]), va='bottom', color='#A5A5A5',fontsize=9)
    ax.annotate(graph_data.at[i,2],xy=(graph_data.at[i,"index"],graph_data.at[i,2]), va='bottom', color='#70AD47',fontsize=9)
    ax.annotate(graph_data.at[i,3],xy=(graph_data.at[i,"index"],graph_data.at[i,3]), va='top', color='#FFC000',fontsize=9)
    ax.annotate(graph_data.at[i,4],xy=(graph_data.at[i,"index"],graph_data.at[i,4]), va='top', color='#ED7D31',fontsize=9)
    ax.annotate(graph_data.at[i,5],xy=(graph_data.at[i,"index"],graph_data.at[i,5]), va='bottom', color='#6600CC',fontsize=9)

ax.yaxis.tick_right()
ax.legend(loc='upper center', bbox_to_anchor=(0.5,1.15), ncol=6)
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

#value display
graph_data=graph_data.reset_index()
for i in range(len(graph_data.index)):
    ax.annotate(graph_data.at[i,0],xy=(graph_data.at[i,"index"],graph_data.at[i,0]), va='bottom', color='#4472C4',fontsize=9)
    ax.annotate(graph_data.at[i,1],xy=(graph_data.at[i,"index"],graph_data.at[i,1]), va='bottom', color='#A5A5A5',fontsize=9)
    ax.annotate(graph_data.at[i,2],xy=(graph_data.at[i,"index"],graph_data.at[i,2]), va='bottom', color='#70AD47',fontsize=9)
    ax.annotate(graph_data.at[i,3],xy=(graph_data.at[i,"index"],graph_data.at[i,3]), va='bottom', color='#FFC000',fontsize=9)
    ax.annotate(graph_data.at[i,4],xy=(graph_data.at[i,"index"],graph_data.at[i,4]), va='bottom', color='#ED7D31',fontsize=9)
    ax.annotate(graph_data.at[i,5],xy=(graph_data.at[i,"index"],graph_data.at[i,5]), va='bottom', color='#6600CC',fontsize=9)
    ax.annotate(graph_data.at[i,6],xy=(graph_data.at[i,"index"],graph_data.at[i,6]), va='bottom', color='#FF66FF',fontsize=9)

ax.yaxis.tick_right()
ax.legend(loc='upper center', bbox_to_anchor=(0.5,1.15), ncol=7)
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

#value display
graph_data=graph_data.reset_index()
for i in range(len(graph_data.index)):
    ax.annotate(graph_data.at[i,0],xy=(graph_data.at[i,"index"],graph_data.at[i,0]), va='bottom', color='#4472C4',fontsize=9)
    ax.annotate(graph_data.at[i,1],xy=(graph_data.at[i,"index"],graph_data.at[i,1]), va='bottom', color='#A5A5A5',fontsize=9)
    ax.annotate(graph_data.at[i,2],xy=(graph_data.at[i,"index"],graph_data.at[i,2]), va='bottom', color='#70AD47',fontsize=9)
    ax.annotate(graph_data.at[i,3],xy=(graph_data.at[i,"index"],graph_data.at[i,3]), va='bottom', color='#FFC000',fontsize=9)
    ax.annotate(graph_data.at[i,4],xy=(graph_data.at[i,"index"],graph_data.at[i,4]), va='bottom', color='#ED7D31',fontsize=9)

ax.yaxis.tick_right()
ax.legend(loc='upper center', bbox_to_anchor=(0.5,1.15), ncol=5)
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

#value display
graph_data=graph_data.reset_index()
for i in range(len(graph_data.index)):
    ax.annotate(graph_data.at[i,0],xy=(graph_data.at[i,"index"],graph_data.at[i,0]), va='bottom', color='#4472C4',fontsize=9)
    ax.annotate(graph_data.at[i,1],xy=(graph_data.at[i,"index"],graph_data.at[i,1]), va='bottom', color='#A5A5A5',fontsize=9)
    ax.annotate(graph_data.at[i,2],xy=(graph_data.at[i,"index"],graph_data.at[i,2]), va='top', color='#70AD47',fontsize=9)
    ax.annotate(graph_data.at[i,3],xy=(graph_data.at[i,"index"],graph_data.at[i,3]), va='bottom', color='#FFC000',fontsize=9)
    ax.annotate(graph_data.at[i,4],xy=(graph_data.at[i,"index"],graph_data.at[i,4]), va='top', color='#ED7D31',fontsize=9)
    ax.annotate(graph_data.at[i,5],xy=(graph_data.at[i,"index"],graph_data.at[i,5]), va='bottom', color='#6600CC',fontsize=9)

ax.yaxis.tick_right()
ax.legend(loc='upper center', bbox_to_anchor=(0.5,1.15), ncol=6)
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

#value display
graph_data=graph_data.reset_index()
for i in range(len(graph_data.index)):
    ax.annotate(graph_data.at[i,0],xy=(graph_data.at[i,"index"],graph_data.at[i,0]), va='bottom', color='#4472C4',fontsize=9)
    ax.annotate(graph_data.at[i,1],xy=(graph_data.at[i,"index"],graph_data.at[i,1]), va='bottom', color='#A5A5A5',fontsize=9)
    ax.annotate(graph_data.at[i,2],xy=(graph_data.at[i,"index"],graph_data.at[i,2]), va='bottom', color='#70AD47',fontsize=9)
    ax.annotate(graph_data.at[i,3],xy=(graph_data.at[i,"index"],graph_data.at[i,3]), va='bottom', color='#FFC000',fontsize=9)
    ax.annotate(graph_data.at[i,4],xy=(graph_data.at[i,"index"],graph_data.at[i,4]), va='bottom', color='#ED7D31',fontsize=9)
    ax.annotate(graph_data.at[i,5],xy=(graph_data.at[i,"index"],graph_data.at[i,5]), va='bottom', color='#6600CC',fontsize=9)
    ax.annotate(graph_data.at[i,6],xy=(graph_data.at[i,"index"],graph_data.at[i,6]), va='bottom', color='#FF66FF',fontsize=9)

ax.yaxis.tick_right()
ax.legend(loc='upper center', bbox_to_anchor=(0.5,1.15), ncol=7)
#그림 저장
plt.tight_layout()
plt.savefig(save_path+file_name)

