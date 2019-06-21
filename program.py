import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
from matplotlib.pyplot import *
import numpy as np
from matplotlib import rc
pd.options.mode.chained_assignment = None  # default='warn'

total_installed_capacity = pd.read_excel('data.xlsx', sheet_name='2013_Total Installed Capacity', index=False)
capacity_additions = pd.read_excel('data.xlsx', sheet_name='Capacity Additions', index=False)
electricity_generation = pd.read_excel('data.xlsx', sheet_name='Electricity Generation', index=False)

technology = ['Coal','Biomass','Large Hydro','Onshore Wind',
                                        'Offshore Wind', 'Utility-scale PV', 'Solar thermal',
                                        'Geothermal', 'Nuclear', 'Natural Gas', 'Oil',
                                        'Marine', 'Storage', 'Small-scale PV']

colors = ['#557f2d','#2d7f5e', '#A6012C', '#FF0042', '#bfb220', '#3fc6bd', '#159e96',
          '#fa7268', '#ffd531', '#ea9494', '#c082f6', '#a4c2f4', '#af8401', '#72af96']

# ––– Add the key which will serve as a reference
total_installed_capacity.insert(2,"Key", total_installed_capacity['Country'] + '+' + total_installed_capacity['Technology'] , True)
capacity_additions.insert(2,"Key", capacity_additions['Country'] + '+' + capacity_additions['Technology'] , True)
electricity_generation.insert(2,"Key", electricity_generation['Country'] + '+' + electricity_generation['Technology'] , True)


"""            Quetion 1             """

total_installed_capacity = pd.merge(total_installed_capacity, capacity_additions, how='left')
total_installed_capacity

for i in range(0,2020-2013):
    total_installed_capacity[2014+i] = total_installed_capacity[2014+i] + total_installed_capacity[2013+i]

total_installed_capacity.to_excel('Q1_solution.xlsx')



"""            Question 2            """

portugal_tech_mix_evolution = total_installed_capacity.loc[total_installed_capacity['Country'] == 'Portugal'].convert_objects(convert_numeric=True)

legend = portugal_tech_mix_evolution['Technology']
legend = legend.reset_index()
legend = legend.drop('index', axis=1)

portugal_tech_mix_evolution = portugal_tech_mix_evolution.reset_index()
portugal_tech_mix_evolution = portugal_tech_mix_evolution.drop(['index','Country','Key','Technology'], axis=1)
portugal_tech_mix_evolution = portugal_tech_mix_evolution.transpose()
portugal_tech_mix_evolution.columns = technology
portugal_tech_mix_evolution


tech_mix = pd.DataFrame(portugal_tech_mix_evolution)
ax = tech_mix.plot(kind='line', figsize=(20,20),color=colors, markersize=22, linewidth=3)
ax.legend(technology, loc='lower center', bbox_to_anchor=(0.5, -0.12), borderaxespad=0.5, prop={'size': 18}, ncol=7);

ax.set_axisbelow(True)
ax.set_facecolor('white')
ax.grid(b=True, which='major', color='#999999', linestyle='-')
ax.minorticks_on()
ax.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
ax.xaxis.set_tick_params(labelsize=20)
ax.yaxis.set_tick_params(labelsize=20)
plt.ylabel('Capacity (MW)', fontsize=18)
plt.xlabel('Year', fontsize=18)
plt.title('Evolution of technology mix in Portugal', fontsize=32)
plt.savefig('Evolution_techmix_Portugal.png')


"""            Question 3           """
transpose = portugal_tech_mix_evolution.transpose()
cumsum = np.cumsum(transpose)
r = [0,1,2,3,4,5,6,7]

# Names of group and bar width
names = ['2013','2014','2015','2016','2017', '2018', '2019', '2020']
barWidth = 0.5
 
f, ax = plt.subplots(figsize=(20,20))
ax.grid(b=True, which='major', color='#999999', linestyle='-')

plt.bar(r, transpose.iloc[0,:], color='#7f6d5f', edgecolor='white', width=barWidth)
for i in range (0, 13):
     plt.bar(r, transpose.iloc[i+1,:], bottom=cumsum.iloc[i,:], color=colors[i], edgecolor='white', width=barWidth)

ax.legend(technology, loc='lower center',bbox_to_anchor=(0.5, -0.12),prop={'size': 14}, ncol=7);
plt.xticks(r, names)
ax.set_facecolor('white')
ax.set_axisbelow(True)
ax.xaxis.set_tick_params(labelsize=20)
ax.yaxis.set_tick_params(labelsize=20)
plt.ylabel('Capacity (MW)', fontsize=18)
plt.xlabel('Year', fontsize=18)
plt.title('Total cumulative capacity', fontsize=32)
plt.savefig('Cumulative_generation_Portugal.png')

"""            Question 4               """

days = 365 # number of days in a year
hours= 24 # number of hours in a day 

portugal_electrcity_output = electricity_generation.loc[electricity_generation['Key'] == 'Portugal+Coal'].convert_objects(convert_numeric=True)
portugal_coal_capactiy = total_installed_capacity.loc[total_installed_capacity['Key'] == 'Portugal+Coal']
portugal_coal_generation = portugal_coal_capactiy
capacity_factor = portugal_coal_capactiy

for i in range(0,2021-2013):
    portugal_coal_generation[2013+i] = (portugal_coal_capactiy[2013+i] * 24 *365)/1000  #Convert from MW to GWh
    capacity_factor[2013+i] = portugal_electrcity_output[2013+i].values / portugal_coal_generation[2013+i].values 

capacity_factor.to_excel('Q4_solution.xlsx')