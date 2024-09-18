from pyomo.opt import SolverFactory
from DCOPF import DCOPFmodel
from main import dataprinter
from tabulate import tabulate
import pandas as pd

file_name = 'UK 29BUSES - 2.xlsx'
all_cost=0
dataprinter(file_name)

model = DCOPFmodel()

solver = 'gurobi'
optimise = SolverFactory(solver)
#
instance = model.create_instance('data.dat')
results = optimise.solve(instance)
print(results)


# print('--------------------------------------------------------')
# print('Optimal_outputs_for_generators')
# Optimal_outputs_for_generators = []
# Optimal_outputs_for_generators.append(['Name of Generator','Scenario','Time Period','Generation(MW)'])
# for t in instance.Time:
#     for g in instance.Generator:
#             Optimal_outputs_for_generators.append([t,g,instance.pG[t,g].value])
# print (tabulate(Optimal_outputs_for_generators))


print('--------------------------------------------------------')
print('The flow of Lines')
The_flow_of_Lines = []
The_flow_of_Lines.append(['Name of Line','Scenario','Time Period','Power Flow(MW)'])
for l in instance.Line:
    for t in instance.Time:
            The_flow_of_Lines.append([l,t,instance.pL[l,t].value])
print (tabulate(The_flow_of_Lines))

# print('--------------------------------------------------------')
# print('Demand')
# Demand = []
# Demand.append(['Name of Demand','Scenario','Time Period','Value of Demand(MW)'])
# for d in instance.Demand:
#     for t in instance.Time:
#             Demand.append([d,t,instance.DT[t,d]])
# print (tabulate(Demand))

a=0
print('--------------------------------------------------------')
print('Cost of Generation')
CoG=[]
CoG.append(['Name of Generator','Scenario','Time Period','Cost','Total Cost(£)'])
for t in instance.Time:
    for g in instance.Generator:
            cost = (instance.pG[t,g].value) * (instance.Cost[g])
            a+=cost
            CoG.append([g,t,cost,a])
print(tabulate(CoG))








cols_Summary = ['Time Period', 'Generation (GW)',  'Demand (GW)','Cost of Generators']
cols_Bus = ['Time Period', 'Name','Angle(degs)']
cols_Generation = ['Time Period', 'Name','Generation Output (GW)']
cols_Line = ['Time Period', 'Name','From', 'To', 'Flow', 'SLmax','increase percentage']



Summary = pd.DataFrame(columns=cols_Summary)
Bus = pd.DataFrame(columns=cols_Bus)
Generation = pd.DataFrame(columns=cols_Generation)
Line = pd.DataFrame(columns=cols_Line)



# Sheet Summary
ind = 0
for t in instance.Time:
        Summary.loc[ind] = pd.Series({'Time Period':t,
                                  'Generation (MW)':sum(instance.pG[t,g].value  for g in instance.Generator),\
                                  'Demand (GW)':sum(instance.DT[t,d] for d in instance.Demand),\
                                  'Cost of Generators':sum(instance.pG[t,g].value * instance.Cost[g] for g in instance.Generator),\
                                  })

        ind += 1


# Sheet Buses
ind = 0
for t in instance.Time:
    for b in instance.Buses:
            Bus.loc[ind] = pd.Series({'Time Period':t,
                                  'Name':b,\
                                  'Angle(degs)':instance.delta[b,t].value,\
                               })
            ind += 1






#Sheet Generator
ind = 0
for t in instance.Time:
    for g in instance.Generator:
            Generation.loc[ind] = pd.Series({'Time Period':t, \
                                     'Name':g,\
                                     'Generation Output (GW)':instance.pG[t,g].value
                                  })
            ind += 1


#Sheet of Lines
ind = 0
for t in instance.Time:
    for l in instance.Line:
            Line.loc[ind] = pd.Series({'Time Period':t, \
                                     'Name':l,\
                                     'From':instance.A[l,1],\
                                     'To':instance.A[l,2],\
                                     'Flow':instance.pL[l,t].value, \
                                    'SLmax':instance.SLmax[l],
                                    'increase percentage':abs(instance.pL[l,t].value/instance.pL[l,2023].value)
                                  })
            ind += 1





Bus = Bus.sort_values(['Name'])
Generation = Generation.sort_values(['Time Period'])
writer = pd.ExcelWriter('results.xlsx', engine ='xlsxwriter')
Summary.to_excel(writer, sheet_name = 'summary',index=False)
Bus.to_excel(writer, sheet_name = 'bus',index=False)
Line.to_excel(writer, sheet_name = 'line',index=False)
Generation.to_excel(writer, sheet_name = 'generator',index=False)
writer.close()




