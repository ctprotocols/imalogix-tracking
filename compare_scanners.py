# -*- coding: utf-8 -*-
"""
Created on Wed Dec  1 09:39:35 2021

@author: EastmanE
"""

import pandas as pd
import datetime as dt
import matplotlib.pyplot as plt
import io 

df = pd.read_csv('allexams.csv')

df.insert(2, 'Month', 0)
for i in df.index:
    if df.loc[i, 'CTDIvol (mGy)'] == ' ':
        df.drop(i, inplace = True)
    elif df.loc[i, 'CTDIvol (mGy)']== '':
        df.drop(i, inplace = True)
    elif (',' in df.loc[i, 'Exam']) | (',' in df.loc[i, 'Description']):
        df.drop(i, inplace = True)
    else:
        df.loc[i, 'Month'] = dt.datetime.strptime(df.loc[i,'Date'], '%m/%d/%Y').month                

                      
    

    
df['CTDIvol (mGy)']= df['CTDIvol (mGy)'].astype(float)

# print(df['CTDIvol (mGy)'])




x=df.groupby(['Exam'])['Scanner'].apply(list)

for name in df.groupby(['Exam']).groups.keys():
    if len(set(x[name])) == 1:
        for i in df.index:
            if df.loc[i, 'Exam'] == name:
                df.drop(i, inplace = True)
            else:
                pass           

examnames = list(set(df['Exam'].tolist()))

with pd.ExcelWriter('outputchart.xlsx') as writer:
    for ename in examnames:

    
        df1 = df.loc[df['Exam'] == ename]
        
        summarygroup = df1.groupby(['Exam', 'Scanner']).agg({'CTDIvol (mGy)': ['median']})
        groups = df1.groupby(['Scanner', 'Month']).agg({'CTDIvol (mGy)': ['median']})
        
        
        summarygroup.to_excel(writer, sheet_name = ename.replace('/', '')[:31])
        groups.to_excel(writer, sheet_name = ename.replace('/', '')[:31], startrow = summarygroup.shape[0]+4)

        scannames = list(set(df1['Scanner'].tolist()))
        
        ws = writer.sheets[ename.replace('/', '')[:31]]
        row = 1
        for sname in scannames:
            chartfilename = io.BytesIO()
    
            df2 = df1.loc[df1['Scanner'] == sname]
    
            boxplot =df2.boxplot(column = 'CTDIvol (mGy)', by = ['Month'], \
                            grid = False, figsize = (10,4), medianprops={'linewidth': 3, 'color': 'red'}, \
                            showfliers=False, return_type = 'dict')
            plt.ylabel('CTDIvol (mGy)')
            plt.xlabel('Month')
            
            plottitle = ename + '- ' + sname 
            plt.title(plottitle)
            plt.suptitle('')
            
            
            plt.savefig(chartfilename, bbox_inches = "tight")
            cell = 'E' + str(row)
            ws.insert_image(cell, '', {'image_data': chartfilename})
            row += 21
            plt.close()
    

# x=df.groupby(['DICOM Protocol Name']).apply(list)
# x=df.groupby(['Exam'])['Scanner'].apply(list)

# for name in df.groupby(['Exam']).groups.keys():
#     if len(set(x[name])) == 1:
#         for i in df.index:
#             if df.loc[i, 'Exam'] == name:
#                 df.drop(i, inplace = True)
#             else:
#                 pass           


# groups = df.groupby(['Exam', 'Scanner', 'Month']).agg({'CTDIvol (mGy)': ['median']})





# groups.to_excel('output.xlsx')

