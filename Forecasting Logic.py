from openpyxl import load_workbook
import xlsxwriter
import openpyxl
from openpyxl.styles import Font
import win32com.client as win32
import pandas as pd
import numpy as np


#     input2.iloc[-1,:-1] = input2.iloc[-1,:-1].apply(lambda x : int(x*(input2.iloc[-1,-1]/input.iloc[-1,-1])))
    



Columnnames = ['Family','SKU','Metrics','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar','Total']

AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'All India',skiprows=5 )
AllIndianew = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'All India (2)',skiprows=5 )

AllIndia = AllIndia.iloc[:,:16]
AllIndianew = AllIndianew.iloc[:,:16]

AllIndia.columns = Columnnames
AllIndianew.columns = Columnnames

AllIndia[['Family','SKU','Metrics']] = AllIndia[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = AllIndia.index[(AllIndia['Family']=='Total')&(AllIndia['Metrics']=='Stock days')]
AllIndia = AllIndia.iloc[:finalrow[0]+1]
skuremovalrowindexes = AllIndia.index[(AllIndia['Family']=='Total')&(AllIndia['SKU'] == 'P Platform Total')]

AllIndianew[['Family','SKU','Metrics']] = AllIndianew[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = AllIndianew.index[(AllIndianew['Family']=='Total')&(AllIndianew['Metrics']=='Stock days')]
AllIndianew = AllIndianew.iloc[:finalrow[0]+1]
skuremovalrowindexes = AllIndianew.index[(AllIndianew['Family']=='Total')&(AllIndianew['SKU'] == 'P Platform Total')]

for i in skuremovalrowindexes:
    AllIndia.at[i,'SKU'] = 'Total'

for i in skuremovalrowindexes:
    AllIndianew.at[i,'SKU'] = 'Total'

AllIndiaRetail = AllIndia.query("Metrics in ['Retail']")   
AllIndianewRetail = AllIndianew.query("Metrics in ['Retail']").copy()
AllIndianewRetail.reset_index(drop=True, inplace=True)

modeltotal = ['Meteor']
models = []
modelfamily = []
for i in AllIndiaRetail['SKU']:
    if 'Total' in i:
        modeltotal.append(i)
for i in range(len(AllIndiaRetail)):
    if 'Total' not in AllIndiaRetail.iloc[i,1]:
        models.append(AllIndiaRetail.iloc[i,0]+"--"+AllIndiaRetail.iloc[i,1])   
for i in AllIndiaRetail['Family']:
    if i not in modelfamily and i != 'Total':
        modelfamily.append(i)     
models.remove('Meteor--Meteor')        
print(models)
print(modelfamily)
# print(type(AllIndianewRetail.columns))
# print(AllIndianewRetail.columns)
if AllIndianewRetail.iloc[-1,-1] != AllIndiaRetail.iloc[-1,-1]:
    print('All India Final value is different')
    totalratio = AllIndianewRetail.iloc[-1,-1]/AllIndiaRetail.iloc[-1,-1]
    # print(AllIndianewRetail.iloc[-1,-1],AllIndiaRetail.iloc[-1,-1])
    summodeltotal = 0
    for i in modeltotal[:-1]:
        # print(AllIndiaRetail.loc[(AllIndiaRetail['SKU']==i),['Total']].values)
        indtotalvalue = round(AllIndiaRetail.loc[(AllIndiaRetail['SKU']==i),['Total']].values[0][0] * totalratio)
        AllIndianewRetail.loc[AllIndianewRetail.SKU == i,'Total'] = indtotalvalue
        summodeltotal += indtotalvalue
    # print(AllIndianewRetail)        
    if summodeltotal != AllIndianewRetail.iloc[-1,-1]:
        if AllIndianewRetail.iloc[-1,-1] - summodeltotal > 0:
            # print(modeltotal[:-1])
            minmodeltotal = AllIndianewRetail[AllIndianewRetail['SKU'].isin(modeltotal[:-1])].copy()
            minmodeltotal['Total'] = minmodeltotal['Total'].astype(int)
            if AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] != 0:
                AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] + (AllIndianewRetail.iloc[-1,-1] - summodeltotal)
            else:
                minval = minmodeltotal.iloc[0,-1]
                minindex = 0
                for i in minmodeltotal['Total'].index:
                    if ((minmodeltotal.iloc[i,-1]< minval) or (minval == 0)) and minmodeltotal.iloc[i,-1] !=0:
                        minindex = i 
                        minval = minmodeltotal.iloc[i,-1]
                AllIndianewRetail.iloc[minindex,-1] = AllIndianewRetail.iloc[minindex,-1] + (AllIndianewRetail.iloc[-1,-1] - summodeltotal)        
        else:
            minmodeltotal = AllIndianewRetail[AllIndianewRetail['SKU'].isin(modeltotal[:-1])].copy()
            minmodeltotal['Total'] = minmodeltotal['Total'].astype(int)
            if AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] != 0:
                AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] - ( summodeltotal - AllIndianewRetail.iloc[-1,-1] )
            else:
                # maxval = minmodeltotal.iloc[0,-1]
                # maxindex = 0
                # for i in minmodeltotal['Total'].index:
                #     if minmodeltotal.iloc[i,-1]> maxval:
                #         maxindex = i 
                # AllIndianewRetail.iloc[maxindex,-1] = AllIndianewRetail.iloc[maxindex,-1] - ( summodeltotal - AllIndianewRetail.iloc[-1,-1] )
                print('max value is 0')
    for f in modelfamily:
        # print(f)
        summodel = 0
        for i in models:
            # print(AllIndiaRetail.loc[(AllIndiaRetail['SKU']==i),['Total']].values)

            if i.startswith(f):
                # print(f,i)
                # print(i.split("--")[0],i.split("--")[1])                
                indtotalvalue = round(AllIndiaRetail.loc[(AllIndiaRetail['Family']==f)&(AllIndiaRetail['SKU']==i.split("--")[1]),['Total']].values[0][0] * totalratio)
                # print(indtotalvalue)
                AllIndianewRetail.loc[(AllIndianewRetail['Family']==f)&(AllIndianewRetail.SKU == i.split("--")[1]),'Total'] = indtotalvalue
                summodel += indtotalvalue 
        # print(AllIndianewRetail)        
        try:           
            print(AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0],summodel)
        except:
            continue            
        if summodel != AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0]:
            if AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0] - summodel > 0:
                minmodeltotal = AllIndianewRetail[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU'].isin([i.split("--")[1] for i in models]))].copy()
                minmodeltotal['Total'] = minmodeltotal['Total'].astype(int)
                # AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] + 1
                # print(minmodeltotal['Total'].idxmin(axis=0)) 
                if AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] != 0:
                    AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] + (AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0] - summodel)
                else:
                    minval = minmodeltotal.iloc[0,-1]
                    minindex = 0
                    for i in minmodeltotal['Total'].index:
                        if ((minmodeltotal.iloc[i,-1]< minval) or (minval == 0)) and minmodeltotal.iloc[i,-1]!=0:
                            minindex = i 
                            minval = minmodeltotal.iloc[i,-1]
                    AllIndianewRetail.iloc[minindex,-1] = AllIndianewRetail.iloc[minindex,-1] + (AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0] - summodel)

            else:  
                minmodeltotal = AllIndianewRetail[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU'].isin([i.split("--")[1] for i in models]))].copy()
                minmodeltotal['Total'] = minmodeltotal['Total'].astype(int)
                # AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] - 1
                # print(minmodeltotal['Total'].idxmax(axis=0)) 
                if AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] != 0:
                    AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] - ( summodel - AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0])
                else:
                    # maxval = minmodeltotal.iloc[0,-1]
                    # maxindex = 0
                    # for i in minmodeltotal['Total'].index:
                    #     if minmodeltotal.iloc[i,-1]> maxval:
                    #         maxindex = i 
                    # AllIndianewRetail.iloc[maxindex,-1] = AllIndianewRetail.iloc[maxindex,-1] - ( summodel - AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0])
                    print('max value is 0')
        # print(round(AllIndianewRetail.loc[AllIndianewRetail['SKU']==i,['Total']]))
    for i in range(len(AllIndianewRetail)):
        iter = 3
        if 'Total' not in AllIndianewRetail.iloc[i,1]:
            for j in AllIndianewRetail.iloc[i,3:-1]:
                indtotalvalue = round(AllIndianewRetail.iloc[i,iter]*totalratio)
                AllIndianewRetail.iloc[i,iter] = indtotalvalue
                iter = iter + 1
        if sum(AllIndianewRetail.iloc[i,3:-1]) > AllIndianewRetail.iloc[i,-1] :
            # print(i,AllIndianewRetail.iloc[i,3:-1].idxmax(axis = 0))
            # if AllIndianewRetail.iloc[i,AllIndianewRetail.iloc[i,3:-1].idxmax(axis = 1)]!=0:
            #     AllIndianewRetail.iloc[i,AllIndianewRetail.iloc[i,3:-1].idxmax(axis = 1)]  = AllIndianewRetail.iloc[i,AllIndianewRetail.iloc[i,3:-1].idxmax(axis = 1)] - (sum(AllIndianewRetail.iloc[i,3:-1]) - AllIndianewRetail.iloc[i,-1])    
            maxindex = 3
            maxval = AllIndianewRetail.iloc[i,maxindex]
            for j in AllIndianewRetail.iloc[i,3:-1].index:
                ji = AllIndianewRetail.columns.get_loc(j)
                if AllIndianewRetail.iloc[i,ji] > maxval:
                    maxindex = ji
            AllIndianewRetail.iloc[i,maxindex] = AllIndianewRetail.iloc[i,maxindex] - (sum(AllIndianewRetail.iloc[i,3:-1]) - AllIndianewRetail.iloc[i,-1])
            # else:
            #     print('max value is 0')       
        else :
            # print(i,AllIndianewRetail.iloc[i,3:-1].idxmax(axis = 1))
            # if AllIndianewRetail.iloc[i,AllIndianewRetail.iloc[i,3:-1].idxmin(axis = 1)]!=0:
            #     AllIndianewRetail.iloc[i,AllIndianewRetail.iloc[i,3:-1].idxmin(axis = 1)]  = AllIndianewRetail.iloc[i,AllIndianewRetail.iloc[i,3:-1].idxmin(axis = 1)] + ( AllIndianewRetail.iloc[i,-1] - sum(AllIndianewRetail.iloc[i,3:-1]) )    
            # else:
            minindex = 3
            minval = AllIndianewRetail.iloc[i,minindex]
            for j in AllIndianewRetail.iloc[i,3:-1].index:
                ji = AllIndianewRetail.columns.get_loc(j)
                if ((AllIndianewRetail.iloc[i,ji] < minval) or (minval == 0)) and AllIndianewRetail.iloc[i,ji] != 0:
                    minindex = ji
                    minval = AllIndianewRetail.iloc[i,minindex]
            AllIndianewRetail.iloc[i,minindex] = AllIndianewRetail.iloc[i,minindex] +  + ( AllIndianewRetail.iloc[i,-1] - sum(AllIndianewRetail.iloc[i,3:-1]) )    
    iter = 3        
    for i in AllIndianewRetail.iloc[:,3:-1]: 
        sumverticalmain = 0
        for j in modelfamily:
            sumvertical = 0
            for k in models:              
                if k.split('--')[0] == j:
                    # print(i,j,k.split('--')[1])
                    # print(int(AllIndianewRetail.loc[(AllIndianewRetail['Family']==j)&(AllIndianewRetail['SKU']==k.split('--')[1]),i] ))
                    sumvertical += int(AllIndianewRetail.loc[(AllIndianewRetail['Family']==j)&(AllIndianewRetail['SKU']==k.split('--')[1]),i] )    
            AllIndianewRetail.loc[(AllIndianewRetail['SKU']==j+ " Total"),i]  = sumvertical 
            sumverticalmain += sumvertical
        AllIndianewRetail.iloc[-1,iter]= sumverticalmain + int(AllIndianewRetail.loc[(AllIndianewRetail['SKU']=='Meteor'),i])
        iter += 1

AllIndiaRetail.reset_index(drop=True, inplace=True)
# print(AllIndiaRetail.iloc[:-1,-1].values)
# print(AllIndianewRetail.iloc[:-1,-1].values)
if np.array_equal(AllIndiaRetail.iloc[:-1,-1].values,AllIndianewRetail.iloc[:-1,-1].values):
    # print('Check is working')
    pass
else:
    Allindvalue = 0
    for f in modelfamily:
        summodel = 0
        if AllIndianewRetail.loc[(AllIndianewRetail['SKU']== f + " Total"),['Total']].values != AllIndiaRetail.loc[(AllIndiaRetail['SKU']== f+ " Total"),['Total']].values:
            for m in models:
                
                if m.startswith(f):  
                    # print(f,m)       
                    indtotalvalue = round(AllIndiaRetail.loc[(AllIndiaRetail['Family']==f)&(AllIndiaRetail['SKU']==m.split("--")[1]),['Total']].values[0][0] * (AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+ " Total"),['Total']].values[0][0]/AllIndiaRetail.loc[(AllIndiaRetail['SKU']==f+ " Total"),['Total']].values[0][0]))
                    # print(indtotalvalue)
                    AllIndianewRetail.loc[(AllIndianewRetail['Family']==f)&(AllIndianewRetail.SKU == m.split("--")[1]),'Total'] = indtotalvalue
                    summodel += indtotalvalue                 
            try:           
                print(AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0],summodel)
            except:
                continue            
            if summodel != AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0]:
                if AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0] - summodel > 0:
                    minmodeltotal = AllIndianewRetail[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU'].isin([i.split("--")[1] for i in models]))].copy()
                    minmodeltotal['Total'] = minmodeltotal['Total'].astype(int)
                    if AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] != 0:
                        AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmin(axis=0),-1] + (AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0] - summodel)
                    else:
                        minval = minmodeltotal.iloc[0,-1]
                        minindex = 0
                        for i in minmodeltotal['Total'].index:
                            if ((minmodeltotal.iloc[i,-1]< minval) or (minval == 0)) and minmodeltotal.iloc[i,-1]!=0:
                                minindex = i 
                                minval = minmodeltotal.iloc[i,-1]
                        AllIndianewRetail.iloc[minindex,-1] = AllIndianewRetail.iloc[minindex,-1] + (AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0] - summodel)

                else:  
                    minmodeltotal = AllIndianewRetail[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU'].isin([i.split("--")[1] for i in models]))].copy()
                    minmodeltotal['Total'] = minmodeltotal['Total'].astype(int)
                    if AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] != 0:
                        AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] = AllIndianewRetail.iloc[minmodeltotal['Total'].idxmax(axis=0),-1] - ( summodel - AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+" Total"),'Total'].values[0])
                    else:
                        print('max value is 0')    
            summodel = 0
            for m in models:
                if m.split('--')[0] == f and 'Total' not in m:
                    indtotalvalue = AllIndianewRetail.loc[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU']==m.split('--')[1]),'Total'].values[0]
                    summodel += indtotalvalue                   
                # if m.startswith(f):  
            AllIndianewRetail.loc[(AllIndianewRetail['SKU']== f + " Total"),['Total']] = summodel                        
        elif AllIndianewRetail.loc[(AllIndianewRetail['SKU']== f + " Total"),['Total']].values == AllIndiaRetail.loc[(AllIndiaRetail['SKU']== f+ " Total"),['Total']].values:
            # print(f)
            summodel = 0
            for m in models:
                if m.split('--')[0] == f and 'Total' not in m:
                    indtotalvalue = AllIndianewRetail.loc[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU']==m.split('--')[1]),'Total'].values[0]
                    print(m.split('--')[1],indtotalvalue)
                    summodel += indtotalvalue                   
                # if m.startswith(f):  
            AllIndianewRetail.loc[(AllIndianewRetail['SKU']== f + " Total"),['Total']] = summodel
        if f != 'Meteor':    
            Allindvalue += AllIndianewRetail.loc[(AllIndianewRetail['SKU']== f + " Total"),'Total'].values[0]
    AllIndianewRetail.iloc[-1,-1] = Allindvalue + AllIndianewRetail.loc[(AllIndianewRetail['SKU']=='Meteor'),'Total'].values[0]
    totalratio = AllIndianewRetail.iloc[-1,-1]/AllIndiaRetail.iloc[-1,-1]
    for i in range(len(AllIndianewRetail)):
        iter = 3
        if 'Total' not in AllIndianewRetail.iloc[i,1]:
            for j in AllIndianewRetail.iloc[i,3:-1]:
                indtotalvalue = round(AllIndianewRetail.iloc[i,iter]*totalratio)
                AllIndianewRetail.iloc[i,iter] = indtotalvalue
                iter = iter + 1
        if sum(AllIndianewRetail.iloc[i,3:-1]) > AllIndianewRetail.iloc[i,-1] :
            maxindex = 3
            maxval = AllIndianewRetail.iloc[i,maxindex]
            for j in AllIndianewRetail.iloc[i,3:-1].index:
                ji = AllIndianewRetail.columns.get_loc(j)
                if AllIndianewRetail.iloc[i,ji] > maxval:
                    maxindex = ji
            AllIndianewRetail.iloc[i,maxindex] = AllIndianewRetail.iloc[i,maxindex] - (sum(AllIndianewRetail.iloc[i,3:-1]) - AllIndianewRetail.iloc[i,-1])
      
        else :

            minindex = 3
            minval = AllIndianewRetail.iloc[i,minindex]
            for j in AllIndianewRetail.iloc[i,3:-1].index:
                ji = AllIndianewRetail.columns.get_loc(j)
                if ((AllIndianewRetail.iloc[i,ji] < minval) or (minval == 0)) and AllIndianewRetail.iloc[i,ji] != 0:
                    minindex = ji
                    minval = AllIndianewRetail.iloc[i,minindex]
            AllIndianewRetail.iloc[i,minindex] = AllIndianewRetail.iloc[i,minindex] +  + ( AllIndianewRetail.iloc[i,-1] - sum(AllIndianewRetail.iloc[i,3:-1]) )    
    iter = 3        
    for i in AllIndianewRetail.iloc[:,3:-1]: 
        sumverticalmain = 0
        for j in modelfamily:
            sumvertical = 0
            for k in models:              
                if k.split('--')[0] == j:
                    sumvertical += int(AllIndianewRetail.loc[(AllIndianewRetail['Family']==j)&(AllIndianewRetail['SKU']==k.split('--')[1]),i] )    
            AllIndianewRetail.loc[(AllIndianewRetail['SKU']==j+ " Total"),i]  = sumvertical 
            sumverticalmain += sumvertical
        AllIndianewRetail.iloc[-1,iter]= sumverticalmain + int(AllIndianewRetail.loc[(AllIndianewRetail['SKU']=='Meteor'),i])
        iter += 1   
if np.array_equal(AllIndianewRetail.iloc[-1,:-1].values , AllIndiaRetail.iloc[-1,:-1].values):
    # print('yes')
    pass
else:
    for i in range(3,len(AllIndianewRetail.columns)-1):
        if AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i] != AllIndiaRetail.iloc[len(AllIndianewRetail)-1,i]:
            totalratio = AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i]/AllIndiaRetail.iloc[len(AllIndianewRetail)-1,i]
            summonth = 0
            for m in models:
                indvalue =  round(AllIndiaRetail.loc[(AllIndiaRetail['Family']==m.split('--')[0])&(AllIndiaRetail['SKU']==m.split('--')[1]),AllIndiaRetail.columns[i]].values[0]*totalratio )
                AllIndianewRetail.loc[(AllIndianewRetail['Family']==m.split('--')[0])&(AllIndianewRetail['SKU']==m.split('--')[1]),AllIndianewRetail.columns[i]]  = indvalue
                summonth += indvalue         
                # print(indvalue,AllIndiaRetail.loc[(AllIndiaRetail['Family']==m.split('--')[0])&(AllIndiaRetail['SKU']==m.split('--')[1]),AllIndiaRetail.columns[i]].values[0])    
            indvalue =  round(AllIndianewRetail.loc[(AllIndianewRetail['SKU']=='Meteor'),AllIndianewRetail.columns[i]].values[0]*totalratio )                             
            AllIndianewRetail.loc[(AllIndianewRetail['SKU']=='Meteor'),AllIndianewRetail.columns[i]] = indvalue
            summonth += indvalue
            if  AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i] > summonth :
                # print(AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i],summonth)
                # print(modeltotal[:-1])
                minmodeltotal = AllIndianewRetail[~AllIndianewRetail['SKU'].isin(modeltotal)].copy()
                minmodeltotal.reset_index(inplace=True)
                # print(minmodeltotal)
                minmodeltotal[minmodeltotal.columns[i+1]] = minmodeltotal[minmodeltotal.columns[i+1]].astype(int)
                # print(minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmin(axis=0),i+1])
                if minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmin(axis=0),i+1] != 0:
                    AllIndianewRetail.iloc[minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmin(axis=0),0],i+1] = minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmin(axis=0),i+1] + ( AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i]  - summonth)
                else:
                    minval = minmodeltotal.iloc[0,i+1]
                    minindex = 0
                    for ind in range(len(minmodeltotal)):
                        if ((minmodeltotal.iloc[ind,i+1]< minval) or (minval == 0)) and minmodeltotal.iloc[ind,i+1] !=0:
                            minindex = minmodeltotal.iloc[ind,0] 
                            minval = minmodeltotal.iloc[ind,i+1]
                    AllIndianewRetail.iloc[minindex,i] = AllIndianewRetail.iloc[minindex,i] + (AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i]  - summonth )                  
            elif  summonth > AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i] :
                # print(AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i],summonth)
                minmodeltotal = AllIndianewRetail[~AllIndianewRetail['SKU'].isin(modeltotal)].copy()
                minmodeltotal.reset_index(inplace=True)
                # print(minmodeltotal)
                minmodeltotal[minmodeltotal.columns[i+1]] = minmodeltotal[minmodeltotal.columns[i+1]].astype(int)
                # print(minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmax(axis=0),i+1])
                if minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmax(axis=0),i+1] != 0:
                    AllIndianewRetail.iloc[minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmax(axis=0),0],i] = minmodeltotal.iloc[minmodeltotal[minmodeltotal.columns[i+1]].idxmax(axis=0),i+1] - ( summonth -AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i]   )
                else:
                    maxval = minmodeltotal.iloc[0,i+1]
                    maxindex = 0
                    for ind in range(len(minmodeltotal)):
                        if ((minmodeltotal.iloc[ind,i+1] > maxval) or (maxval == 0)) and minmodeltotal.iloc[ind,i+1] !=0:
                            maxindex = minmodeltotal.iloc[ind,0] 
                            maxval = minmodeltotal.iloc[ind,i+1]
                    AllIndianewRetail.iloc[maxindex,i] = AllIndianewRetail.iloc[maxindex,i] - ( summonth -AllIndianewRetail.iloc[len(AllIndianewRetail)-1,i]    )   
        for f in modelfamily:
            summodel = 0
            for m in models:
                if m.startswith(f) and 'Total' not in m:
                    summodel += AllIndianewRetail.loc[(AllIndianewRetail['Family']==f)&(AllIndianewRetail['SKU']==m.split('--')[1]),AllIndianewRetail.columns[i]].values[0]  
            AllIndianewRetail.loc[(AllIndianewRetail['SKU']==f+ ' Total'),AllIndianewRetail.columns[i]] = summodel

    for row in range(len(AllIndianewRetail)):
        sumrow = 0
        for col in range(3,len(AllIndianewRetail.columns)-1):
            sumrow += AllIndianewRetail.iloc[row,col]     
        AllIndianewRetail.iloc[row,-1] = sumrow

N = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'North Zone TTL',skiprows=5 )        
S = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'South Zone TTL',skiprows=5 )        
W = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'West Zone TTL',skiprows=5 )        
E = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'East Zone TTL',skiprows=5 )        
# print(W)

N = N.iloc[:,:16]
S = S.iloc[:,:16]
W = W.iloc[:,:16]
E = E.iloc[:,:16]

N.columns = Columnnames
W.columns = Columnnames
S.columns = Columnnames
E.columns = Columnnames

N[['Family','SKU','Metrics']] = N[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N.index[(N['Family']=='Total')&(N['Metrics']=='Stock days')]
N = N.iloc[:finalrow[0]+1]
skuremovalrowindexes = N.index[(N['Family']=='Total')&(N['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N.at[i,'SKU'] = 'Total'

S[['Family','SKU','Metrics']] = S[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S.index[(S['Family']=='Total')&(S['Metrics']=='Stock days')]
S = S.iloc[:finalrow[0]+1]
skuremovalrowindexes = S.index[(S['Family']=='Total')&(S['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S.at[i,'SKU'] = 'Total'

W[['Family','SKU','Metrics']] = W[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = W.index[(W['Family']=='Total')&(W['Metrics']=='Stock days')]
W = W.iloc[:finalrow[0]+1]
skuremovalrowindexes = W.index[(W['Family']=='Total')&(W['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    W.at[i,'SKU'] = 'Total'

E[['Family','SKU','Metrics']] = E[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = E.index[(E['Family']=='Total')&(E['Metrics']=='Stock days')]
E = E.iloc[:finalrow[0]+1]
skuremovalrowindexes = E.index[(E['Family']=='Total')&(E['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    E.at[i,'SKU'] = 'Total'

NRetail = N.query("Metrics in ['Retail']")  
SRetail = S.query("Metrics in ['Retail']")
WRetail = W.query("Metrics in ['Retail']")
ERetail = E.query("Metrics in ['Retail']")


NRetailnew = N.query("Metrics in ['Retail']").copy()
NRetailnew.reset_index(drop=True, inplace=True)
SRetailnew = S.query("Metrics in ['Retail']").copy()
SRetailnew.reset_index(drop=True, inplace=True)
WRetailnew = W.query("Metrics in ['Retail']").copy()
WRetailnew.reset_index(drop=True, inplace=True)
ERetailnew = E.query("Metrics in ['Retail']").copy()
ERetailnew.reset_index(drop=True, inplace=True)
zones = [NRetailnew,ERetailnew,WRetailnew,SRetailnew]
for row in range(len(AllIndianewRetail)):
    for col in range(3,len(AllIndianewRetail.columns)-1):
        if 'Total' not in AllIndianewRetail.iloc[row,1] :
            try:
                NRetailnew.iloc[row,col] = round(NRetail.iloc[row,col]*(AllIndianewRetail.iloc[row,col]/AllIndiaRetail.iloc[row,col]))
                ERetailnew.iloc[row,col] = round(ERetail.iloc[row,col]*(AllIndianewRetail.iloc[row,col]/AllIndiaRetail.iloc[row,col]))
                WRetailnew.iloc[row,col] = round(WRetail.iloc[row,col]*(AllIndianewRetail.iloc[row,col]/AllIndiaRetail.iloc[row,col]))
                SRetailnew.iloc[row,col] = round(SRetail.iloc[row,col]*(AllIndianewRetail.iloc[row,col]/AllIndiaRetail.iloc[row,col]))
                # print('Value changed for : ',row,col)
            except:
                continue    
            sumzone = NRetailnew.iloc[row,col] + ERetailnew.iloc[row,col] + WRetailnew.iloc[row,col] + SRetailnew.iloc[row,col]
            if sumzone > AllIndianewRetail.iloc[row,col]:
                maxvalue = NRetailnew.iloc[row,col]
                maxindex = 0
                for zone in range(len(zones)):
                    if zones[zone].iloc[row,col] > maxvalue:
                        maxvalue = zones[zone].iloc[row,col]
                        maxindex = zone
                zones[maxindex].iloc[row,col] = maxvalue - (sumzone - AllIndianewRetail.iloc[row,col])

            elif AllIndianewRetail.iloc[row,col] > sumzone:
                minval =   NRetailnew.iloc[row,col]
                minindex = 0
                for zone in range(len(zones)):
                    if (zones[zone].iloc[row,col] < minval or minval ==0) and zones[zone].iloc[row,col] != 0:
                        minval = zones[zone].iloc[row,col]
                        minindex = zone 
                zones[minindex].iloc[row,col] = minval + (AllIndianewRetail.iloc[row,col] - sumzone) 

for i in range(3,len(AllIndianewRetail.columns)-1):    
    for f in modelfamily:
        summodelN = 0
        summodelE = 0
        summodelW = 0
        summodelS = 0            
        for m in models:
            if m.startswith(f) and 'Total' not in m:
                summodelN += NRetailnew.loc[(NRetailnew['Family']==f)&(NRetailnew['SKU']==m.split('--')[1]),NRetailnew.columns[i]].values[0]  
                summodelE += ERetailnew.loc[(ERetailnew['Family']==f)&(ERetailnew['SKU']==m.split('--')[1]),ERetailnew.columns[i]].values[0]  
                summodelW += WRetailnew.loc[(WRetailnew['Family']==f)&(WRetailnew['SKU']==m.split('--')[1]),WRetailnew.columns[i]].values[0]  
                summodelS += SRetailnew.loc[(SRetailnew['Family']==f)&(SRetailnew['SKU']==m.split('--')[1]),SRetailnew.columns[i]].values[0]  

        SRetailnew.loc[(SRetailnew['SKU']==f+ ' Total'),SRetailnew.columns[i]] = summodelS
        WRetailnew.loc[(WRetailnew['SKU']==f+ ' Total'),WRetailnew.columns[i]] = summodelW
        ERetailnew.loc[(ERetailnew['SKU']==f+ ' Total'),ERetailnew.columns[i]] = summodelE
        NRetailnew.loc[(NRetailnew['SKU']==f+ ' Total'),NRetailnew.columns[i]] = summodelN
    # print(SRetailnew.loc[(SRetailnew['SKU'].isin(modeltotal[:-1])),SRetailnew.columns[i]])    
    SRetailnew.iloc[-1,i] = sum(SRetailnew.loc[(SRetailnew['SKU'].isin(modeltotal[:-1])),SRetailnew.columns[i]])
    WRetailnew.iloc[-1,i] = sum(WRetailnew.loc[(WRetailnew['SKU'].isin(modeltotal[:-1])),WRetailnew.columns[i]])
    ERetailnew.iloc[-1,i] = sum(ERetailnew.loc[(ERetailnew['SKU'].isin(modeltotal[:-1])),ERetailnew.columns[i]])
    NRetailnew.iloc[-1,i] = sum(NRetailnew.loc[(NRetailnew['SKU'].isin(modeltotal[:-1])),NRetailnew.columns[i]])

for row in range(len(AllIndianewRetail)):
    SumNrow = 0
    SumErow = 0
    SumSrow = 0
    SumWrow = 0
    for col in range(3,len(AllIndianewRetail.columns)-1):
        SumNrow += NRetailnew.iloc[row,col]  
        SumErow += ERetailnew.iloc[row,col]    
        SumSrow += SRetailnew.iloc[row,col]     
        SumWrow += WRetailnew.iloc[row,col]    
    NRetailnew.iloc[row,-1] = SumNrow
    SRetailnew.iloc[row,-1] = SumSrow
    ERetailnew.iloc[row,-1] = SumErow
    WRetailnew.iloc[row,-1] = SumWrow    


with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    NRetailnew.to_excel(writer, sheet_name='North') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    WRetailnew.to_excel(writer, sheet_name='West') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    ERetailnew.to_excel(writer, sheet_name='East') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    SRetailnew.to_excel(writer, sheet_name='South')             
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    AllIndianewRetail.to_excel(writer, sheet_name='AllIndiaResult') 



N1 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N1',skiprows=5 )        
N2 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N2',skiprows=5 )        
N3 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N3',skiprows=5 )        
N4 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N4',skiprows=5 )        
N5 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N5',skiprows=5 )        
N6 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N6',skiprows=5 )        
N7 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N7',skiprows=5 )        
N8 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'N8',skiprows=5 )        

# print(W)


N1 = N1.iloc[:,:16]
N2 = N2.iloc[:,:16]
N3 = N3.iloc[:,:16]
N4 = N4.iloc[:,:16]
N5 = N5.iloc[:,:16]
N6 = N6.iloc[:,:16]
N7 = N7.iloc[:,:16]
N8 = N8.iloc[:,:16]



N1.columns = Columnnames
N2.columns = Columnnames
N3.columns = Columnnames
N4.columns = Columnnames
N5.columns = Columnnames
N6.columns = Columnnames
N7.columns = Columnnames
N8.columns = Columnnames

# print(N1)
# print(N3)

N1[['Family','SKU','Metrics']] = N1[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N1.index[(N1['Family']=='Total')&(N1['Metrics']=='Stock days')]
N1 = N1.iloc[:finalrow[0]+1]
skuremovalrowindexes = N1.index[(N1['Family']=='Total')&(N1['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N1.at[i,'SKU'] = 'Total'

N2[['Family','SKU','Metrics']] = N2[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N2.index[(N2['Family']=='Total')&(N2['Metrics']=='Stock days')]
N2 = N2.iloc[:finalrow[0]+1]
skuremovalrowindexes = N2.index[(N2['Family']=='Total')&(N2['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N2.at[i,'SKU'] = 'Total'

N3[['Family','SKU','Metrics']] = N3[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N3.index[(N3['Family']=='Total')&(N3['Metrics']=='Stock days')]
N3 = N3.iloc[:finalrow[0]+1]
skuremovalrowindexes = N3.index[(N3['Family']=='Total')&(N3['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N3.at[i,'SKU'] = 'Total'

N4[['Family','SKU','Metrics']] = N4[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N4.index[(N4['Family']=='Total')&(N4['Metrics']=='Stock days')]
N4 = N4.iloc[:finalrow[0]+1]
skuremovalrowindexes = N4.index[(N4['Family']=='Total')&(N4['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N4.at[i,'SKU'] = 'Total'

N5[['Family','SKU','Metrics']] = N5[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N5.index[(N5['Family']=='Total')&(N5['Metrics']=='Stock days')]
N5 = N5.iloc[:finalrow[0]+1]
skuremovalrowindexes = N5.index[(N5['Family']=='Total')&(N5['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N5.at[i,'SKU'] = 'Total'

N6[['Family','SKU','Metrics']] = N6[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N6.index[(N6['Family']=='Total')&(N6['Metrics']=='Stock days')]
N6 = N6.iloc[:finalrow[0]+1]
skuremovalrowindexes = N6.index[(N6['Family']=='Total')&(N6['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N6.at[i,'SKU'] = 'Total'

N7[['Family','SKU','Metrics']] = N7[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N7.index[(N7['Family']=='Total')&(N7['Metrics']=='Stock days')]
N7 = N7.iloc[:finalrow[0]+1]
skuremovalrowindexes = N7.index[(N7['Family']=='Total')&(N7['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N7.at[i,'SKU'] = 'Total'

N8[['Family','SKU','Metrics']] = N8[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = N8.index[(N8['Family']=='Total')&(N8['Metrics']=='Stock days')]
N8 = N8.iloc[:finalrow[0]+1]
skuremovalrowindexes = N8.index[(N8['Family']=='Total')&(N8['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    N8.at[i,'SKU'] = 'Total'  

N1 = N1.fillna(0)
N2 = N2.fillna(0)
N3 = N3.fillna(0)
N4 = N4.fillna(0)
N5 = N5.fillna(0)
N6 = N6.fillna(0)
N7 = N7.fillna(0)
N8 = N8.fillna(0)

N1Retail = N1.query("Metrics in ['Retail']")  
N2Retail = N2.query("Metrics in ['Retail']")
N3Retail = N3.query("Metrics in ['Retail']")
N4Retail = N4.query("Metrics in ['Retail']")
N5Retail = N5.query("Metrics in ['Retail']")  
N6Retail = N6.query("Metrics in ['Retail']")  
N7Retail = N7.query("Metrics in ['Retail']")  
N8Retail = N8.query("Metrics in ['Retail']")  


N1Retailnew = N1.query("Metrics in ['Retail']").copy()
N1Retailnew.reset_index(drop=True, inplace=True)
N2Retailnew = N2.query("Metrics in ['Retail']").copy()
N2Retailnew.reset_index(drop=True, inplace=True)
N3Retailnew = N3.query("Metrics in ['Retail']").copy()
N3Retailnew.reset_index(drop=True, inplace=True)
N4Retailnew = N4.query("Metrics in ['Retail']").copy()
N4Retailnew.reset_index(drop=True, inplace=True)
N5Retailnew = N5.query("Metrics in ['Retail']").copy()
N5Retailnew.reset_index(drop=True, inplace=True)
N6Retailnew = N6.query("Metrics in ['Retail']").copy()
N6Retailnew.reset_index(drop=True, inplace=True)
N7Retailnew = N7.query("Metrics in ['Retail']").copy()
N7Retailnew.reset_index(drop=True, inplace=True)
N8Retailnew = N8.query("Metrics in ['Retail']").copy()
N8Retailnew.reset_index(drop=True, inplace=True)
# print(N1Retailnew)
# print(N3Retailnew)
NRegiones = [N1Retailnew,N2Retailnew,N3Retailnew,N4Retailnew,N5Retailnew,N6Retailnew,N7Retailnew,N8Retailnew]

for row in range(len(NRetailnew)):
    for col in range(3,len(NRetailnew.columns)-1):
        if 'Total' not in NRetailnew.iloc[row,1] :
            try:
                N1Retailnew.iloc[row,col] = round(N1Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N2Retailnew.iloc[row,col] = round(N2Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N3Retailnew.iloc[row,col] = round(N3Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N4Retailnew.iloc[row,col] = round(N4Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N5Retailnew.iloc[row,col] = round(N5Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N6Retailnew.iloc[row,col] = round(N6Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N7Retailnew.iloc[row,col] = round(N7Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))
                N8Retailnew.iloc[row,col] = round(N8Retail.iloc[row,col]*(NRetailnew.iloc[row,col]/NRetail.iloc[row,col]))                
                # print('Value changed for : ',row,col)
            except:
                continue    
            sumzone = N1Retailnew.iloc[row,col] + N2Retailnew.iloc[row,col] + N3Retailnew.iloc[row,col] + N4Retailnew.iloc[row,col] + N5Retailnew.iloc[row,col] + N6Retailnew.iloc[row,col] + N7Retailnew.iloc[row,col] + N8Retailnew.iloc[row,col]
            # print(sumzone,NRetailnew.iloc[row,col])
            if sumzone > NRetailnew.iloc[row,col]:
                maxvalue = N1Retailnew.iloc[row,col]
                maxindex = 0
                for zone in range(len(NRegiones)):
                    if NRegiones[zone].iloc[row,col] > maxvalue:
                        maxvalue = NRegiones[zone].iloc[row,col]
                        maxindex = zone
                NRegiones[maxindex].iloc[row,col] = maxvalue - (sumzone - NRetailnew.iloc[row,col])
                # print(NRegiones[maxindex].iloc[row,col],maxvalue,(sumzone - NRetailnew.iloc[row,col]))

            elif NRetailnew.iloc[row,col] > sumzone:
                minval =   NRetailnew.iloc[row,col]
                minindex = 0
                for zone in range(len(NRegiones)):
                    if (NRegiones[zone].iloc[row,col] < minval or minval ==0) and NRegiones[zone].iloc[row,col] != 0:
                        minval = NRegiones[zone].iloc[row,col]
                        minindex = zone 
                NRegiones[minindex].iloc[row,col] = minval + (NRetailnew.iloc[row,col] - sumzone) 
                # print(NRegiones[minindex].iloc[row,col],minval,(NRetailnew.iloc[row,col] - sumzone) )
                

for i in range(3,len(NRetailnew.columns)-1):    
    for f in modelfamily:
        summodelN1 = 0
        summodelN2 = 0
        summodelN3 = 0
        summodelN4 = 0            
        summodelN5 = 0            
        summodelN6 = 0            
        summodelN7 = 0            
        summodelN8 = 0            

        for m in models:
            if m.startswith(f) and 'Total' not in m:
                summodelN1 += N1Retailnew.loc[(N1Retailnew['Family']==f)&(N1Retailnew['SKU']==m.split('--')[1]),N1Retailnew.columns[i]].values[0]  
                summodelN2 += N2Retailnew.loc[(N2Retailnew['Family']==f)&(N2Retailnew['SKU']==m.split('--')[1]),N2Retailnew.columns[i]].values[0]  
                summodelN3 += N3Retailnew.loc[(N3Retailnew['Family']==f)&(N3Retailnew['SKU']==m.split('--')[1]),N3Retailnew.columns[i]].values[0]  
                summodelN4 += N4Retailnew.loc[(N4Retailnew['Family']==f)&(N4Retailnew['SKU']==m.split('--')[1]),N4Retailnew.columns[i]].values[0]  
                summodelN5 += N5Retailnew.loc[(N5Retailnew['Family']==f)&(N5Retailnew['SKU']==m.split('--')[1]),N5Retailnew.columns[i]].values[0]
                summodelN6 += N6Retailnew.loc[(N6Retailnew['Family']==f)&(N6Retailnew['SKU']==m.split('--')[1]),N6Retailnew.columns[i]].values[0]
                summodelN7 += N7Retailnew.loc[(N7Retailnew['Family']==f)&(N7Retailnew['SKU']==m.split('--')[1]),N7Retailnew.columns[i]].values[0]
                summodelN8 += N8Retailnew.loc[(N8Retailnew['Family']==f)&(N8Retailnew['SKU']==m.split('--')[1]),N8Retailnew.columns[i]].values[0]                                                
        N1Retailnew.loc[(N1Retailnew['SKU']==f+ ' Total'),N1Retailnew.columns[i]] = summodelN1
        N2Retailnew.loc[(N2Retailnew['SKU']==f+ ' Total'),N2Retailnew.columns[i]] = summodelN2
        N3Retailnew.loc[(N3Retailnew['SKU']==f+ ' Total'),N3Retailnew.columns[i]] = summodelN3
        N4Retailnew.loc[(N4Retailnew['SKU']==f+ ' Total'),N4Retailnew.columns[i]] = summodelN4
        N5Retailnew.loc[(N5Retailnew['SKU']==f+ ' Total'),N5Retailnew.columns[i]] = summodelN5
        N6Retailnew.loc[(N6Retailnew['SKU']==f+ ' Total'),N6Retailnew.columns[i]] = summodelN6
        N7Retailnew.loc[(N7Retailnew['SKU']==f+ ' Total'),N7Retailnew.columns[i]] = summodelN7
        N8Retailnew.loc[(N8Retailnew['SKU']==f+ ' Total'),N8Retailnew.columns[i]] = summodelN8                        
            # print(SRetailnew.loc[(SRetailnew['SKU'].isin(modeltotal[:-1])),SRetailnew.columns[i]])    
    N1Retailnew.iloc[-1,i] = sum(N1Retailnew.loc[(N1Retailnew['SKU'].isin(modeltotal[:-1])),N1Retailnew.columns[i]])
    N2Retailnew.iloc[-1,i] = sum(N2Retailnew.loc[(N2Retailnew['SKU'].isin(modeltotal[:-1])),N2Retailnew.columns[i]])
    N3Retailnew.iloc[-1,i] = sum(N3Retailnew.loc[(N3Retailnew['SKU'].isin(modeltotal[:-1])),N3Retailnew.columns[i]])
    N4Retailnew.iloc[-1,i] = sum(N4Retailnew.loc[(N4Retailnew['SKU'].isin(modeltotal[:-1])),N4Retailnew.columns[i]])
    N5Retailnew.iloc[-1,i] = sum(N5Retailnew.loc[(N5Retailnew['SKU'].isin(modeltotal[:-1])),N5Retailnew.columns[i]])
    N6Retailnew.iloc[-1,i] = sum(N6Retailnew.loc[(N6Retailnew['SKU'].isin(modeltotal[:-1])),N6Retailnew.columns[i]])
    N7Retailnew.iloc[-1,i] = sum(N7Retailnew.loc[(N7Retailnew['SKU'].isin(modeltotal[:-1])),N7Retailnew.columns[i]])
    N8Retailnew.iloc[-1,i] = sum(N8Retailnew.loc[(N8Retailnew['SKU'].isin(modeltotal[:-1])),N8Retailnew.columns[i]])


for row in range(len(NRetailnew)):
    SumN1row = 0
    SumN2row = 0
    SumN3row = 0
    SumN4row = 0
    SumN5row = 0
    SumN6row = 0
    SumN7row = 0
    SumN8row = 0    
    for col in range(3,len(NRetailnew.columns)-1):
        SumN1row += N1Retailnew.iloc[row,col]  
        SumN2row += N2Retailnew.iloc[row,col]    
        SumN3row += N3Retailnew.iloc[row,col]     
        SumN4row += N4Retailnew.iloc[row,col]    
        SumN5row += N5Retailnew.iloc[row,col]  
        SumN6row += N6Retailnew.iloc[row,col]    
        SumN7row += N7Retailnew.iloc[row,col]     
        SumN8row += N8Retailnew.iloc[row,col]         
    N1Retailnew.iloc[row,-1] = SumN1row
    N2Retailnew.iloc[row,-1] = SumN2row
    N3Retailnew.iloc[row,-1] = SumN3row
    N4Retailnew.iloc[row,-1] = SumN4row    
    N5Retailnew.iloc[row,-1] = SumN5row
    N6Retailnew.iloc[row,-1] = SumN6row
    N7Retailnew.iloc[row,-1] = SumN7row
    N8Retailnew.iloc[row,-1] = SumN8row     

with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N1Retailnew.to_excel(writer, sheet_name='N1') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N2Retailnew.to_excel(writer, sheet_name='N2') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N3Retailnew.to_excel(writer, sheet_name='N3') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N4Retailnew.to_excel(writer, sheet_name='N4')             
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N5Retailnew.to_excel(writer, sheet_name='N5') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N6Retailnew.to_excel(writer, sheet_name='N6') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N7Retailnew.to_excel(writer, sheet_name='N7') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    N8Retailnew.to_excel(writer, sheet_name='N8')     
     


E1 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'E1',skiprows=5 )        
E2 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'E2',skiprows=5 )        
E3 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'E3',skiprows=5 )        
E4 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'E4',skiprows=5 )        
E5 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'E5',skiprows=5 )        

# print(W)


E1 = E1.iloc[:,:16]
E2 = E2.iloc[:,:16]
E3 = E3.iloc[:,:16]
E4 = E4.iloc[:,:16]
E5 = E5.iloc[:,:16]




E1.columns = Columnnames
E2.columns = Columnnames
E3.columns = Columnnames
E4.columns = Columnnames
E5.columns = Columnnames


# print(E1)
# print(E3)

E1[['Family','SKU','Metrics']] = E1[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = E1.index[(E1['Family']=='Total')&(E1['Metrics']=='Stock days')]
E1 = E1.iloc[:finalrow[0]+1]
skuremovalrowindexes = E1.index[(E1['Family']=='Total')&(E1['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    E1.at[i,'SKU'] = 'Total'

E2[['Family','SKU','Metrics']] = E2[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = E2.index[(E2['Family']=='Total')&(E2['Metrics']=='Stock days')]
E2 = E2.iloc[:finalrow[0]+1]
skuremovalrowindexes = E2.index[(E2['Family']=='Total')&(E2['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    E2.at[i,'SKU'] = 'Total'

E3[['Family','SKU','Metrics']] = E3[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = E3.index[(E3['Family']=='Total')&(E3['Metrics']=='Stock days')]
E3 = E3.iloc[:finalrow[0]+1]
skuremovalrowindexes = E3.index[(E3['Family']=='Total')&(E3['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    E3.at[i,'SKU'] = 'Total'

E4[['Family','SKU','Metrics']] = E4[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = E4.index[(E4['Family']=='Total')&(E4['Metrics']=='Stock days')]
E4 = E4.iloc[:finalrow[0]+1]
skuremovalrowindexes = E4.index[(E4['Family']=='Total')&(E4['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    E4.at[i,'SKU'] = 'Total'

E5[['Family','SKU','Metrics']] = E5[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = E5.index[(E5['Family']=='Total')&(E5['Metrics']=='Stock days')]
E5 = E5.iloc[:finalrow[0]+1]
skuremovalrowindexes = E5.index[(E5['Family']=='Total')&(E5['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    E5.at[i,'SKU'] = 'Total'

E1 = E1.fillna(0)
E2 = E2.fillna(0)
E3 = E3.fillna(0)
E4 = E4.fillna(0)
E5 = E5.fillna(0)


E1Retail = E1.query("Metrics in ['Retail']")  
E2Retail = E2.query("Metrics in ['Retail']")
E3Retail = E3.query("Metrics in ['Retail']")
E4Retail = E4.query("Metrics in ['Retail']")
E5Retail = E5.query("Metrics in ['Retail']")  
  


E1Retailnew = E1.query("Metrics in ['Retail']").copy()
E1Retailnew.reset_index(drop=True, inplace=True)
E2Retailnew = E2.query("Metrics in ['Retail']").copy()
E2Retailnew.reset_index(drop=True, inplace=True)
E3Retailnew = E3.query("Metrics in ['Retail']").copy()
E3Retailnew.reset_index(drop=True, inplace=True)
E4Retailnew = E4.query("Metrics in ['Retail']").copy()
E4Retailnew.reset_index(drop=True, inplace=True)
E5Retailnew = E5.query("Metrics in ['Retail']").copy()
E5Retailnew.reset_index(drop=True, inplace=True)

# print(E1Retailnew)
# print(E3Retailnew)
ERegiones = [E1Retailnew,E2Retailnew,E3Retailnew,E4Retailnew,E5Retailnew]

for row in range(len(ERetailnew)):
    for col in range(3,len(ERetailnew.columns)-1):
        if 'Total' not in ERetailnew.iloc[row,1] :
            try:
                E1Retailnew.iloc[row,col] = round(E1Retail.iloc[row,col]*(ERetailnew.iloc[row,col]/ERetail.iloc[row,col]))
                E2Retailnew.iloc[row,col] = round(E2Retail.iloc[row,col]*(ERetailnew.iloc[row,col]/ERetail.iloc[row,col]))
                E3Retailnew.iloc[row,col] = round(E3Retail.iloc[row,col]*(ERetailnew.iloc[row,col]/ERetail.iloc[row,col]))
                E4Retailnew.iloc[row,col] = round(E4Retail.iloc[row,col]*(ERetailnew.iloc[row,col]/ERetail.iloc[row,col]))
                E5Retailnew.iloc[row,col] = round(E5Retail.iloc[row,col]*(ERetailnew.iloc[row,col]/ERetail.iloc[row,col]))
              
                # print('Value changed for : ',row,col)
            except:
                continue    
            sumzone = E1Retailnew.iloc[row,col] + E2Retailnew.iloc[row,col] + E3Retailnew.iloc[row,col] + E4Retailnew.iloc[row,col] + E5Retailnew.iloc[row,col] 
            # print(sumzone,ERetailnew.iloc[row,col])
            if sumzone > ERetailnew.iloc[row,col]:
                maxvalue = E1Retailnew.iloc[row,col]
                maxindex = 0
                for zone in range(len(ERegiones)):
                    if ERegiones[zone].iloc[row,col] > maxvalue:
                        maxvalue = ERegiones[zone].iloc[row,col]
                        maxindex = zone
                ERegiones[maxindex].iloc[row,col] = maxvalue - (sumzone - ERetailnew.iloc[row,col])
                # print(ERegiones[maxindex].iloc[row,col],maxvalue,(sumzone - ERetailnew.iloc[row,col]))

            elif ERetailnew.iloc[row,col] > sumzone:
                minval =   ERetailnew.iloc[row,col]
                minindex = 0
                for zone in range(len(ERegiones)):
                    if (ERegiones[zone].iloc[row,col] < minval or minval ==0) and ERegiones[zone].iloc[row,col] != 0:
                        minval = ERegiones[zone].iloc[row,col]
                        minindex = zone 
                ERegiones[minindex].iloc[row,col] = minval + (ERetailnew.iloc[row,col] - sumzone) 
                # print(ERegiones[minindex].iloc[row,col],minval,(ERetailnew.iloc[row,col] - sumzone) )
                

for i in range(3,len(ERetailnew.columns)-1):    
    for f in modelfamily:
        summodelE1 = 0
        summodelE2 = 0
        summodelE3 = 0
        summodelE4 = 0            
        summodelE5 = 0            
          

        for m in models:
            if m.startswith(f) and 'Total' not in m:
                summodelE1 += E1Retailnew.loc[(E1Retailnew['Family']==f)&(E1Retailnew['SKU']==m.split('--')[1]),E1Retailnew.columns[i]].values[0]  
                summodelE2 += E2Retailnew.loc[(E2Retailnew['Family']==f)&(E2Retailnew['SKU']==m.split('--')[1]),E2Retailnew.columns[i]].values[0]  
                summodelE3 += E3Retailnew.loc[(E3Retailnew['Family']==f)&(E3Retailnew['SKU']==m.split('--')[1]),E3Retailnew.columns[i]].values[0]  
                summodelE4 += E4Retailnew.loc[(E4Retailnew['Family']==f)&(E4Retailnew['SKU']==m.split('--')[1]),E4Retailnew.columns[i]].values[0]  
                summodelE5 += E5Retailnew.loc[(E5Retailnew['Family']==f)&(E5Retailnew['SKU']==m.split('--')[1]),E5Retailnew.columns[i]].values[0]
                                               
        E1Retailnew.loc[(E1Retailnew['SKU']==f+ ' Total'),E1Retailnew.columns[i]] = summodelE1
        E2Retailnew.loc[(E2Retailnew['SKU']==f+ ' Total'),E2Retailnew.columns[i]] = summodelE2
        E3Retailnew.loc[(E3Retailnew['SKU']==f+ ' Total'),E3Retailnew.columns[i]] = summodelE3
        E4Retailnew.loc[(E4Retailnew['SKU']==f+ ' Total'),E4Retailnew.columns[i]] = summodelE4
        E5Retailnew.loc[(E5Retailnew['SKU']==f+ ' Total'),E5Retailnew.columns[i]] = summodelE5
                      
            # print(SRetailnew.loc[(SRetailnew['SKU'].isin(modeltotal[:-1])),SRetailnew.columns[i]])    
    E1Retailnew.iloc[-1,i] = sum(E1Retailnew.loc[(E1Retailnew['SKU'].isin(modeltotal[:-1])),E1Retailnew.columns[i]])
    E2Retailnew.iloc[-1,i] = sum(E2Retailnew.loc[(E2Retailnew['SKU'].isin(modeltotal[:-1])),E2Retailnew.columns[i]])
    E3Retailnew.iloc[-1,i] = sum(E3Retailnew.loc[(E3Retailnew['SKU'].isin(modeltotal[:-1])),E3Retailnew.columns[i]])
    E4Retailnew.iloc[-1,i] = sum(E4Retailnew.loc[(E4Retailnew['SKU'].isin(modeltotal[:-1])),E4Retailnew.columns[i]])
    E5Retailnew.iloc[-1,i] = sum(E5Retailnew.loc[(E5Retailnew['SKU'].isin(modeltotal[:-1])),E5Retailnew.columns[i]])



for row in range(len(ERetailnew)):
    SumE1row = 0
    SumE2row = 0
    SumE3row = 0
    SumE4row = 0
    SumE5row = 0
  
    for col in range(3,len(ERetailnew.columns)-1):
        SumE1row += E1Retailnew.iloc[row,col]  
        SumE2row += E2Retailnew.iloc[row,col]    
        SumE3row += E3Retailnew.iloc[row,col]     
        SumE4row += E4Retailnew.iloc[row,col]    
        SumE5row += E5Retailnew.iloc[row,col]  
        
    E1Retailnew.iloc[row,-1] = SumE1row
    E2Retailnew.iloc[row,-1] = SumE2row
    E3Retailnew.iloc[row,-1] = SumE3row
    E4Retailnew.iloc[row,-1] = SumE4row    
    E5Retailnew.iloc[row,-1] = SumE5row
    

with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    E1Retailnew.to_excel(writer, sheet_name='E1') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    E2Retailnew.to_excel(writer, sheet_name='E2') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    E3Retailnew.to_excel(writer, sheet_name='E3') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    E4Retailnew.to_excel(writer, sheet_name='E4')             
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    E5Retailnew.to_excel(writer, sheet_name='E5') 




S1 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S1',skiprows=5 )        
S2 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S2',skiprows=5 )        
S3 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S3',skiprows=5 )        
S4 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S4',skiprows=5 )        
S5 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S5',skiprows=5 )        
S6 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S6',skiprows=5 )        
S7 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'S7',skiprows=5 )        

# print(W)


S1 = S1.iloc[:,:16]
S2 = S2.iloc[:,:16]
S3 = S3.iloc[:,:16]
S4 = S4.iloc[:,:16]
S5 = S5.iloc[:,:16]
S6 = S6.iloc[:,:16]
S7 = S7.iloc[:,:16]



S1.columns = Columnnames
S2.columns = Columnnames
S3.columns = Columnnames
S4.columns = Columnnames
S5.columns = Columnnames
S6.columns = Columnnames
S7.columns = Columnnames

# print(S1)
# print(S3)

S1[['Family','SKU','Metrics']] = S1[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S1.index[(S1['Family']=='Total')&(S1['Metrics']=='Stock days')]
S1 = S1.iloc[:finalrow[0]+1]
skuremovalrowindexes = S1.index[(S1['Family']=='Total')&(S1['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S1.at[i,'SKU'] = 'Total'

S2[['Family','SKU','Metrics']] = S2[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S2.index[(S2['Family']=='Total')&(S2['Metrics']=='Stock days')]
S2 = S2.iloc[:finalrow[0]+1]
skuremovalrowindexes = S2.index[(S2['Family']=='Total')&(S2['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S2.at[i,'SKU'] = 'Total'

S3[['Family','SKU','Metrics']] = S3[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S3.index[(S3['Family']=='Total')&(S3['Metrics']=='Stock days')]
S3 = S3.iloc[:finalrow[0]+1]
skuremovalrowindexes = S3.index[(S3['Family']=='Total')&(S3['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S3.at[i,'SKU'] = 'Total'

S4[['Family','SKU','Metrics']] = S4[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S4.index[(S4['Family']=='Total')&(S4['Metrics']=='Stock days')]
S4 = S4.iloc[:finalrow[0]+1]
skuremovalrowindexes = S4.index[(S4['Family']=='Total')&(S4['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S4.at[i,'SKU'] = 'Total'

S5[['Family','SKU','Metrics']] = S5[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S5.index[(S5['Family']=='Total')&(S5['Metrics']=='Stock days')]
S5 = S5.iloc[:finalrow[0]+1]
skuremovalrowindexes = S5.index[(S5['Family']=='Total')&(S5['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S5.at[i,'SKU'] = 'Total'

S6[['Family','SKU','Metrics']] = S6[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S6.index[(S6['Family']=='Total')&(S6['Metrics']=='Stock days')]
S6 = S6.iloc[:finalrow[0]+1]
skuremovalrowindexes = S6.index[(S6['Family']=='Total')&(S6['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S6.at[i,'SKU'] = 'Total'

S7[['Family','SKU','Metrics']] = S7[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = S7.index[(S7['Family']=='Total')&(S7['Metrics']=='Stock days')]
S7 = S7.iloc[:finalrow[0]+1]
skuremovalrowindexes = S7.index[(S7['Family']=='Total')&(S7['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    S7.at[i,'SKU'] = 'Total'



S1 = S1.fillna(0)
S2 = S2.fillna(0)
S3 = S3.fillna(0)
S4 = S4.fillna(0)
S5 = S5.fillna(0)
S6 = S6.fillna(0)
S7 = S7.fillna(0)

S1Retail = S1.query("Metrics in ['Retail']")  
S2Retail = S2.query("Metrics in ['Retail']")
S3Retail = S3.query("Metrics in ['Retail']")
S4Retail = S4.query("Metrics in ['Retail']")
S5Retail = S5.query("Metrics in ['Retail']")  
S6Retail = S6.query("Metrics in ['Retail']")  
S7Retail = S7.query("Metrics in ['Retail']")  


S1Retailnew = S1.query("Metrics in ['Retail']").copy()
S1Retailnew.reset_index(drop=True, inplace=True)
S2Retailnew = S2.query("Metrics in ['Retail']").copy()
S2Retailnew.reset_index(drop=True, inplace=True)
S3Retailnew = S3.query("Metrics in ['Retail']").copy()
S3Retailnew.reset_index(drop=True, inplace=True)
S4Retailnew = S4.query("Metrics in ['Retail']").copy()
S4Retailnew.reset_index(drop=True, inplace=True)
S5Retailnew = S5.query("Metrics in ['Retail']").copy()
S5Retailnew.reset_index(drop=True, inplace=True)
S6Retailnew = S6.query("Metrics in ['Retail']").copy()
S6Retailnew.reset_index(drop=True, inplace=True)
S7Retailnew = S7.query("Metrics in ['Retail']").copy()
S7Retailnew.reset_index(drop=True, inplace=True)

# print(S1Retailnew)
# print(S3Retailnew)
SRegiones = [S1Retailnew,S2Retailnew,S3Retailnew,S4Retailnew,S5Retailnew,S6Retailnew,S7Retailnew]

for row in range(len(SRetailnew)):
    for col in range(3,len(SRetailnew.columns)-1):
        if 'Total' not in SRetailnew.iloc[row,1] :
            try:
                S1Retailnew.iloc[row,col] = round(S1Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                S2Retailnew.iloc[row,col] = round(S2Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                S3Retailnew.iloc[row,col] = round(S3Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                S4Retailnew.iloc[row,col] = round(S4Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                S5Retailnew.iloc[row,col] = round(S5Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                S6Retailnew.iloc[row,col] = round(S6Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                S7Retailnew.iloc[row,col] = round(S7Retail.iloc[row,col]*(SRetailnew.iloc[row,col]/SRetail.iloc[row,col]))
                # print('Value changed for : ',row,col)
            except:
                continue    
            sumzone = S1Retailnew.iloc[row,col] + S2Retailnew.iloc[row,col] + S3Retailnew.iloc[row,col] + S4Retailnew.iloc[row,col] + S5Retailnew.iloc[row,col] + S6Retailnew.iloc[row,col] + S7Retailnew.iloc[row,col] 
            # print(sumzone,SRetailnew.iloc[row,col])
            if sumzone > SRetailnew.iloc[row,col]:
                maxvalue = S1Retailnew.iloc[row,col]
                maxindex = 0
                for zone in range(len(SRegiones)):
                    if SRegiones[zone].iloc[row,col] > maxvalue:
                        maxvalue = SRegiones[zone].iloc[row,col]
                        maxindex = zone
                SRegiones[maxindex].iloc[row,col] = maxvalue - (sumzone - SRetailnew.iloc[row,col])
                # print(SRegiones[maxindex].iloc[row,col],maxvalue,(sumzone - SRetailnew.iloc[row,col]))

            elif SRetailnew.iloc[row,col] > sumzone:
                minval =   SRetailnew.iloc[row,col]
                minindex = 0
                for zone in range(len(SRegiones)):
                    if (SRegiones[zone].iloc[row,col] < minval or minval ==0) and SRegiones[zone].iloc[row,col] != 0:
                        minval = SRegiones[zone].iloc[row,col]
                        minindex = zone 
                SRegiones[minindex].iloc[row,col] = minval + (SRetailnew.iloc[row,col] - sumzone) 
                # print(SRegiones[minindex].iloc[row,col],minval,(SRetailnew.iloc[row,col] - sumzone) )
                

for i in range(3,len(SRetailnew.columns)-1):    
    for f in modelfamily:
        summodelS1 = 0
        summodelS2 = 0
        summodelS3 = 0
        summodelS4 = 0            
        summodelS5 = 0            
        summodelS6 = 0            
        summodelS7 = 0            

        for m in models:
            if m.startswith(f) and 'Total' not in m:
                summodelS1 += S1Retailnew.loc[(S1Retailnew['Family']==f)&(S1Retailnew['SKU']==m.split('--')[1]),S1Retailnew.columns[i]].values[0]  
                summodelS2 += S2Retailnew.loc[(S2Retailnew['Family']==f)&(S2Retailnew['SKU']==m.split('--')[1]),S2Retailnew.columns[i]].values[0]  
                summodelS3 += S3Retailnew.loc[(S3Retailnew['Family']==f)&(S3Retailnew['SKU']==m.split('--')[1]),S3Retailnew.columns[i]].values[0]  
                summodelS4 += S4Retailnew.loc[(S4Retailnew['Family']==f)&(S4Retailnew['SKU']==m.split('--')[1]),S4Retailnew.columns[i]].values[0]  
                summodelS5 += S5Retailnew.loc[(S5Retailnew['Family']==f)&(S5Retailnew['SKU']==m.split('--')[1]),S5Retailnew.columns[i]].values[0]
                summodelS6 += S6Retailnew.loc[(S6Retailnew['Family']==f)&(S6Retailnew['SKU']==m.split('--')[1]),S6Retailnew.columns[i]].values[0]
                summodelS7 += S7Retailnew.loc[(S7Retailnew['Family']==f)&(S7Retailnew['SKU']==m.split('--')[1]),S7Retailnew.columns[i]].values[0]
        S1Retailnew.loc[(S1Retailnew['SKU']==f+ ' Total'),S1Retailnew.columns[i]] = summodelS1
        S2Retailnew.loc[(S2Retailnew['SKU']==f+ ' Total'),S2Retailnew.columns[i]] = summodelS2
        S3Retailnew.loc[(S3Retailnew['SKU']==f+ ' Total'),S3Retailnew.columns[i]] = summodelS3
        S4Retailnew.loc[(S4Retailnew['SKU']==f+ ' Total'),S4Retailnew.columns[i]] = summodelS4
        S5Retailnew.loc[(S5Retailnew['SKU']==f+ ' Total'),S5Retailnew.columns[i]] = summodelS5
        S6Retailnew.loc[(S6Retailnew['SKU']==f+ ' Total'),S6Retailnew.columns[i]] = summodelS6
        S7Retailnew.loc[(S7Retailnew['SKU']==f+ ' Total'),S7Retailnew.columns[i]] = summodelS7
            # print(SRetailnew.loc[(SRetailnew['SKU'].isin(modeltotal[:-1])),SRetailnew.columns[i]])    
    S1Retailnew.iloc[-1,i] = sum(S1Retailnew.loc[(S1Retailnew['SKU'].isin(modeltotal[:-1])),S1Retailnew.columns[i]])
    S2Retailnew.iloc[-1,i] = sum(S2Retailnew.loc[(S2Retailnew['SKU'].isin(modeltotal[:-1])),S2Retailnew.columns[i]])
    S3Retailnew.iloc[-1,i] = sum(S3Retailnew.loc[(S3Retailnew['SKU'].isin(modeltotal[:-1])),S3Retailnew.columns[i]])
    S4Retailnew.iloc[-1,i] = sum(S4Retailnew.loc[(S4Retailnew['SKU'].isin(modeltotal[:-1])),S4Retailnew.columns[i]])
    S5Retailnew.iloc[-1,i] = sum(S5Retailnew.loc[(S5Retailnew['SKU'].isin(modeltotal[:-1])),S5Retailnew.columns[i]])
    S6Retailnew.iloc[-1,i] = sum(S6Retailnew.loc[(S6Retailnew['SKU'].isin(modeltotal[:-1])),S6Retailnew.columns[i]])
    S7Retailnew.iloc[-1,i] = sum(S7Retailnew.loc[(S7Retailnew['SKU'].isin(modeltotal[:-1])),S7Retailnew.columns[i]])


for row in range(len(SRetailnew)):
    SumS1row = 0
    SumS2row = 0
    SumS3row = 0
    SumS4row = 0
    SumS5row = 0
    SumS6row = 0
    SumS7row = 0
    for col in range(3,len(SRetailnew.columns)-1):
        SumS1row += S1Retailnew.iloc[row,col]  
        SumS2row += S2Retailnew.iloc[row,col]    
        SumS3row += S3Retailnew.iloc[row,col]     
        SumS4row += S4Retailnew.iloc[row,col]    
        SumS5row += S5Retailnew.iloc[row,col]  
        SumS6row += S6Retailnew.iloc[row,col]    
        SumS7row += S7Retailnew.iloc[row,col]     
    S1Retailnew.iloc[row,-1] = SumS1row
    S2Retailnew.iloc[row,-1] = SumS2row
    S3Retailnew.iloc[row,-1] = SumS3row
    S4Retailnew.iloc[row,-1] = SumS4row    
    S5Retailnew.iloc[row,-1] = SumS5row
    S6Retailnew.iloc[row,-1] = SumS6row
    S7Retailnew.iloc[row,-1] = SumS7row

with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S1Retailnew.to_excel(writer, sheet_name='S1') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S2Retailnew.to_excel(writer, sheet_name='S2') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S3Retailnew.to_excel(writer, sheet_name='S3') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S4Retailnew.to_excel(writer, sheet_name='S4')             
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S5Retailnew.to_excel(writer, sheet_name='S5') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S6Retailnew.to_excel(writer, sheet_name='S6') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    S7Retailnew.to_excel(writer, sheet_name='S7') 



W1 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'W1',skiprows=5 )        
W2 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'W2',skiprows=5 )        
W3 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'W3',skiprows=5 )        
W4 = AllIndia = pd.read_excel(r"C:\Users\abhishekkd\Downloads\Copy of Field Confidence (ALL Zone Consolidated) - Final.xlsx",sheet_name = 'W4',skiprows=5 )        

# print(W)


W1 = W1.iloc[:,:16]
W2 = W2.iloc[:,:16]
W3 = W3.iloc[:,:16]
W4 = W4.iloc[:,:16]



W1.columns = Columnnames
W2.columns = Columnnames
W3.columns = Columnnames
W4.columns = Columnnames

# print(W1)
# print(W3)

W1[['Family','SKU','Metrics']] = W1[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = W1.index[(W1['Family']=='Total')&(W1['Metrics']=='Stock days')]
W1 = W1.iloc[:finalrow[0]+1]
skuremovalrowindexes = W1.index[(W1['Family']=='Total')&(W1['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    W1.at[i,'SKU'] = 'Total'

W2[['Family','SKU','Metrics']] = W2[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = W2.index[(W2['Family']=='Total')&(W2['Metrics']=='Stock days')]
W2 = W2.iloc[:finalrow[0]+1]
skuremovalrowindexes = W2.index[(W2['Family']=='Total')&(W2['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    W2.at[i,'SKU'] = 'Total'

W3[['Family','SKU','Metrics']] = W3[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = W3.index[(W3['Family']=='Total')&(W3['Metrics']=='Stock days')]
W3 = W3.iloc[:finalrow[0]+1]
skuremovalrowindexes = W3.index[(W3['Family']=='Total')&(W3['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    W3.at[i,'SKU'] = 'Total'

W4[['Family','SKU','Metrics']] = W4[['Family','SKU','Metrics']].fillna(method = 'ffill')
finalrow = W4.index[(W4['Family']=='Total')&(W4['Metrics']=='Stock days')]
W4 = W4.iloc[:finalrow[0]+1]
skuremovalrowindexes = W4.index[(W4['Family']=='Total')&(W4['SKU'] == 'P Platform Total')]
for i in skuremovalrowindexes:
    W4.at[i,'SKU'] = 'Total'


W1 = W1.fillna(0)
W2 = W2.fillna(0)
W3 = W3.fillna(0)
W4 = W4.fillna(0)


W1Retail = W1.query("Metrics in ['Retail']")  
W2Retail = W2.query("Metrics in ['Retail']")
W3Retail = W3.query("Metrics in ['Retail']")
W4Retail = W4.query("Metrics in ['Retail']")



W1Retailnew = W1.query("Metrics in ['Retail']").copy()
W1Retailnew.reset_index(drop=True, inplace=True)
W2Retailnew = W2.query("Metrics in ['Retail']").copy()
W2Retailnew.reset_index(drop=True, inplace=True)
W3Retailnew = W3.query("Metrics in ['Retail']").copy()
W3Retailnew.reset_index(drop=True, inplace=True)
W4Retailnew = W4.query("Metrics in ['Retail']").copy()
W4Retailnew.reset_index(drop=True, inplace=True)

# print(W1Retailnew)
# print(W3Retailnew)
WRegiones = [W1Retailnew,W2Retailnew,W3Retailnew,W4Retailnew]

for row in range(len(WRetailnew)):
    for col in range(3,len(WRetailnew.columns)-1):
        if 'Total' not in WRetailnew.iloc[row,1] :
            try:
                W1Retailnew.iloc[row,col] = round(W1Retail.iloc[row,col]*(WRetailnew.iloc[row,col]/WRetail.iloc[row,col]))
                W2Retailnew.iloc[row,col] = round(W2Retail.iloc[row,col]*(WRetailnew.iloc[row,col]/WRetail.iloc[row,col]))
                W3Retailnew.iloc[row,col] = round(W3Retail.iloc[row,col]*(WRetailnew.iloc[row,col]/WRetail.iloc[row,col]))
                W4Retailnew.iloc[row,col] = round(W4Retail.iloc[row,col]*(WRetailnew.iloc[row,col]/WRetail.iloc[row,col]))

            except:
                continue    
            sumzone = W1Retailnew.iloc[row,col] + W2Retailnew.iloc[row,col] + W3Retailnew.iloc[row,col] + W4Retailnew.iloc[row,col]
            # print(sumzone,WRetailnew.iloc[row,col])
            if sumzone > WRetailnew.iloc[row,col]:
                maxvalue = W1Retailnew.iloc[row,col]
                maxindex = 0
                for zone in range(len(WRegiones)):
                    if WRegiones[zone].iloc[row,col] > maxvalue:
                        maxvalue = WRegiones[zone].iloc[row,col]
                        maxindex = zone
                WRegiones[maxindex].iloc[row,col] = maxvalue - (sumzone - WRetailnew.iloc[row,col])
                # print(WRegiones[maxindex].iloc[row,col],maxvalue,(sumzone - WRetailnew.iloc[row,col]))

            elif WRetailnew.iloc[row,col] > sumzone:
                minval =   WRetailnew.iloc[row,col]
                minindex = 0
                for zone in range(len(WRegiones)):
                    if (WRegiones[zone].iloc[row,col] < minval or minval ==0) and WRegiones[zone].iloc[row,col] != 0:
                        minval = WRegiones[zone].iloc[row,col]
                        minindex = zone 
                WRegiones[minindex].iloc[row,col] = minval + (WRetailnew.iloc[row,col] - sumzone) 
                # print(WRegiones[minindex].iloc[row,col],minval,(WRetailnew.iloc[row,col] - sumzone) )
                

for i in range(3,len(WRetailnew.columns)-1):    
    for f in modelfamily:
        summodelW1 = 0
        summodelW2 = 0
        summodelW3 = 0
        summodelW4 = 0            
          

        for m in models:
            if m.startswith(f) and 'Total' not in m:
                summodelW1 += W1Retailnew.loc[(W1Retailnew['Family']==f)&(W1Retailnew['SKU']==m.split('--')[1]),W1Retailnew.columns[i]].values[0]  
                summodelW2 += W2Retailnew.loc[(W2Retailnew['Family']==f)&(W2Retailnew['SKU']==m.split('--')[1]),W2Retailnew.columns[i]].values[0]  
                summodelW3 += W3Retailnew.loc[(W3Retailnew['Family']==f)&(W3Retailnew['SKU']==m.split('--')[1]),W3Retailnew.columns[i]].values[0]  
                summodelW4 += W4Retailnew.loc[(W4Retailnew['Family']==f)&(W4Retailnew['SKU']==m.split('--')[1]),W4Retailnew.columns[i]].values[0]  
                                              
        W1Retailnew.loc[(W1Retailnew['SKU']==f+ ' Total'),W1Retailnew.columns[i]] = summodelW1
        W2Retailnew.loc[(W2Retailnew['SKU']==f+ ' Total'),W2Retailnew.columns[i]] = summodelW2
        W3Retailnew.loc[(W3Retailnew['SKU']==f+ ' Total'),W3Retailnew.columns[i]] = summodelW3
        W4Retailnew.loc[(W4Retailnew['SKU']==f+ ' Total'),W4Retailnew.columns[i]] = summodelW4
                    
            # print(WRetailnew.loc[(WRetailnew['SKU'].isin(modeltotal[:-1])),WRetailnew.columns[i]])    
    W1Retailnew.iloc[-1,i] = sum(W1Retailnew.loc[(W1Retailnew['SKU'].isin(modeltotal[:-1])),W1Retailnew.columns[i]])
    W2Retailnew.iloc[-1,i] = sum(W2Retailnew.loc[(W2Retailnew['SKU'].isin(modeltotal[:-1])),W2Retailnew.columns[i]])
    W3Retailnew.iloc[-1,i] = sum(W3Retailnew.loc[(W3Retailnew['SKU'].isin(modeltotal[:-1])),W3Retailnew.columns[i]])
    W4Retailnew.iloc[-1,i] = sum(W4Retailnew.loc[(W4Retailnew['SKU'].isin(modeltotal[:-1])),W4Retailnew.columns[i]])

for row in range(len(WRetailnew)):
    SumW1row = 0
    SumW2row = 0
    SumW3row = 0
    SumW4row = 0
  
    for col in range(3,len(WRetailnew.columns)-1):
        SumW1row += W1Retailnew.iloc[row,col]  
        SumW2row += W2Retailnew.iloc[row,col]    
        SumW3row += W3Retailnew.iloc[row,col]     
        SumW4row += W4Retailnew.iloc[row,col]    
        
    W1Retailnew.iloc[row,-1] = SumW1row
    W2Retailnew.iloc[row,-1] = SumW2row
    W3Retailnew.iloc[row,-1] = SumW3row
    W4Retailnew.iloc[row,-1] = SumW4row    


with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    W1Retailnew.to_excel(writer, sheet_name='W1') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    W2Retailnew.to_excel(writer, sheet_name='W2') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    W3Retailnew.to_excel(writer, sheet_name='W3') 
with pd.ExcelWriter(r"C:\Users\abhishekkd\Documents\SBP Test.xlsx", engine='openpyxl', mode='a',if_sheet_exists='replace') as writer:  
    W4Retailnew.to_excel(writer, sheet_name='W4')             
