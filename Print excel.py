from pyomo.opt import SolverStatus, TerminationCondition
from tabulate import tabulate
import pandas as pd
import math
import sys





cols_Summary = ['Time Period', 'Generation (MW)', 'Wind Generation (MW)', 'Demand (MW)','Cost']
cols_Bus = ['Time Period', 'Name' 'Angle(degs)']
cols_Generator = ['Time', 'Name', 'Output']
cols_Wind = ['Time', 'Name', 'Output']
cols_Line = ['Time', 'Name', 'From', 'To', 'Flow', 'Slmax']



summary = pd.DataFrame(columns=cols_Summary)
bus = pd.DataFrame(columns=cols_Bus)
demand = pd.DataFrame(columns=cols_Generator)
wind = pd.DataFrame(columns=cols_Wind)
generation = pd.DataFrame(columns=cols_Line)

summary.loc[0] = pd.Series(
    {'Time Period': sum(instance.pG[g].value for g in instance.G), \
     'Generation (MW)': sum(instance.pW[w].value for w in instance.WIND),\
     'Wind Generation (MW)': sum(instance.PD[d] for d in instance.D), \
     'Demand (MW)':
     'Cost': })


bus = bus.sort_values(['Name'])
generation = generation.sort_values(['Name'])
demand = demand.sort_values(['Name'])
writer = pd.ExcelWriter('results.xlsx', engine ='xlsxwriter')
summary.to_excel(writer, sheet_name = 'summary',index=False)
bus.to_excel(writer, sheet_name = 'bus',index=False)
demand.to_excel(writer, sheet_name = 'demand',index=False)
generation.to_excel(writer, sheet_name = 'generator',index=False)
wind.to_excel(writer, sheet_name = 'wind',index=False)
writer.close()