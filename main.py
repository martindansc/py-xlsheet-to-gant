import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Patch
import sys

xl = pd.ExcelFile('data/' + sys.argv[1] +'.xlsx')
sheet_name = xl.sheet_names[0] 
df = xl.parse(sheet_name)

df = df[::-1]
df = df.reset_index()

# project start date
proj_start = df.Start.min()# number of days from project start to task start
df['start_num'] = df.Start# number of days from project start to end of tasks
df['end_num'] = df.End-proj_start# days between start and end of each task
df['days_start_to_end'] = df.end_num - df.start_num

##### PLOT #####
fig, (ax, ax1) = plt.subplots(2, figsize=(16,6), gridspec_kw={'height_ratios':[6, 1]}, facecolor='#36454F')
ax.set_facecolor('#36454F')
ax1.set_facecolor('#36454F')
# bars
ax.barh(df.Task, df.days_start_to_end, left=df.start_num, color=df.Color, alpha=1)

for idx, row in df.iterrows():
    ax.text(row.start_num-0.1, idx, row.Task, va='center', ha='right', alpha=0.8, color='w')

# grid lines
ax.set_axisbelow(True)
ax.xaxis.grid(color='k', linestyle='dashed', alpha=0.4, which='both')

# ticks
xticks = np.arange(0, df.end_num.max()+1, 3)
#xticks_labels = pd.date_range(proj_start, end=df.End.max()).strftime("%m/%d")
xticks_minor = np.arange(0, df.end_num.max()+1, 1)
ax.set_xticks(xticks)
ax.set_xticks(xticks_minor, minor=True)
# ax.set_xticklabels(xticks_labels[::3], color='w')
ax.tick_params(axis='x', colors='w')
ax.set_yticks([])

plt.setp([ax.get_xticklines()], color='w')

# align x axis
ax.set_xlim(0, df.end_num.max() + 1)

# remove spines
ax.spines['right'].set_visible(False)
ax.spines['left'].set_visible(False)
ax.spines['left'].set_position(('outward', 10))
ax.spines['top'].set_visible(False)
ax.spines['bottom'].set_color('w')

plt.suptitle(sheet_name, color='w')

##### LEGENDS #####
legend_elements = []
departments = df.Department.unique()
print(departments)
for dep_name in reversed(departments):
    dep = df.loc[df['Department'] == dep_name]
    legend_elements.append(Patch(facecolor=dep.iloc[0]['Color'], label=dep_name))


legend = ax1.legend(handles=legend_elements, loc='upper center', ncol=5, frameon=False)
plt.setp(legend.get_texts(), color='w')

# clean second axis
ax1.spines['right'].set_visible(False)
ax1.spines['left'].set_visible(False)
ax1.spines['top'].set_visible(False)
ax1.spines['bottom'].set_visible(False)
ax1.set_xticks([])
ax1.set_yticks([])

plt.savefig('output/' + sheet_name + '.png', facecolor='#36454F')
plt.show()
