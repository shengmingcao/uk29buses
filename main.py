
import pandas as pd
import datetime

def dataprinter(file_name):
    df_bus=pd.read_excel(file_name, sheet_name='bus')
    df_demand=pd.read_excel(file_name, sheet_name='demand')
    df_branch=pd.read_excel(file_name, sheet_name='branch')
    df_generator=pd.read_excel(file_name, sheet_name='generator')
    df_tdemand=pd.read_excel(file_name, sheet_name='tdemand')
    df_tgenerator = pd.read_excel(file_name, sheet_name='tgenerator')
    df_time = pd.read_excel(file_name, sheet_name='time')




    file = open('data.dat','w')
    file.write('#this is the data which is used to test DCOPF with time coupling')
    file.write('#_Author_Shengming Cao\n')
    file.write('#Time stamp: ' + str(datetime.datetime.now()) + '\n')

    ##the construction of sets##
    file.write('set Buses :=\n')
    for i in df_bus.index:
        file.write(str(df_bus['name'][i])+'\n')
    file.write(';\n')

    file.write('set Demand :=\n')
    for i in df_demand.index:
        file.write(str(df_demand['name'][i])+'\n')
    file.write(';\n')

    file.write('set Line :=\n')
    for i in df_branch.index:
        file.write(str(df_branch['name'][i])+'\n')
    file.write(';\n')

    file.write('set Generator :=\n')
    for i in df_generator.index:
        file.write(str(df_generator['name'][i])+'\n')
    file.write(';\n')

    file.write('set Time :=\n')
    for i in df_time.index:
        file.write(str(df_time['Time period'][i])+'\n')
    file.write(';\n')




    #MAPPING FOR GENERATOR AND BUSES#
    file.write('set GenBus:=\n')
    for i in df_generator.index:
        file.write(str(df_generator['busname'][i])+' '+str(df_generator['name'][i])+'\n')
    file.write(';\n')

    #MAPPING FOR  DEMAND AND BUSES#

    file.write('set DemBus:=\n')
    for i in df_demand.index:
        file.write(str(df_demand['busname'][i])+ ' ' +str(df_demand['name'][i])+'\n')
    file.write(';\n')

    file.write('set LE:=\n 1 \n 2;\n')

    # parameters#

    # the capacity of line
    file.write('param SLmax:=\n')
    for i in df_branch.index:
        file.write(str(df_branch['name'][i]) + ' ' + str(df_branch['s_nomgw'][i])+ '\n')
    file.write(';\n')



    #demand with time coupling
    lst_demand = df_demand['name'].tolist()
    file.write('param DT:=\n')
    for i in df_tdemand.index:
        for demand in lst_demand:
            file.write(str(df_tdemand['Time period'][i])+ ' '+ str(demand)+ ' ' +str(df_tdemand[demand][i])+'\n')
    file.write(';\n')

    #generator with time coupling
    lst_generator = df_generator['name'].tolist()
    file.write('param GT:=\n')
    for i in df_tgenerator.index:
        for generator in lst_generator:
            file.write(str(df_tgenerator['Time period'][i])+ ' '+ str(generator)+ ' ' +str(df_tgenerator[generator][i])+'\n')
    file.write(';\n')

    file.write('param BL:=\n')
    x=df_branch['x']
    for i in df_branch.index:
        file.write(str(df_branch['name'][i]) + ' ' + str(1/x[i]) + '\n')
    file.write(';\n')

    # ---param defining system topolgy---
    file.write('param A:=\n')
    for i in df_branch.index:
        file.write(str(df_branch["name"][i]) + " " + "1" + " " + str(df_branch["from_busname"][i]) + "\n")
        file.write(str(df_branch["name"][i]) + " " + "2" + " " + str(df_branch["to_busname"][i]) + "\n")
    file.write(';\n')

    file.write('param Cost:=\n')
    for i in df_generator.index:
        file.write(str(df_generator['name'][i])+ ' ' + str(df_generator['cost'][i]) + '\n')
    file.write(';\n')





