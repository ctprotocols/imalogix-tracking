# -*- coding: utf-8 -*-
"""
Created on Fri Oct 30 13:25:11 2020

@author: EastmanE
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import calendar
import io 

path = r'Z:\Emi\Code\Imalogix_tracking_code\CTDI\CTDI_Tracking_Enter_Data.xlsx'
plotpath = r'Z:\Emi\Code\Imalogix_tracking_code\CTDI\CTDI_Tracking_Output.xlsx'

headers = ['Protocol name', '25 percentile', 'Median', '75 percentile']
summary_df = pd.DataFrame(columns= headers)

xl = pd.ExcelFile(path)
sheet_names = xl.sheet_names

with pd.ExcelWriter(plotpath, engine = 'xlsxwriter') as writer:
    summary_df.to_excel(writer, sheet_name = 'Summary', index = False)
    for name in sheet_names:
        print(name)
        bpfilename = io.BytesIO()
        chartfilename = io.BytesIO()
        df = pd.read_excel(path, sheet_name = name)

        tempdf = pd.DataFrame(data = [[name, np.percentile(df['CTDIvol (mGy)'].tolist(), 25), np.median(df['CTDIvol (mGy)'].tolist()), np.percentile(df['CTDIvol (mGy)'].tolist(), 75)]], columns = headers)
        summary_df = pd.concat([summary_df, tempdf])
        if name == 'Head':
            df = df[['CTDIvol (mGy)', 'Date', 'Description']]
        else:
            df = df[['CTDIvol (mGy)', 'Date']]

        if name == 'CCTA':
            df = df.loc[df['CTDIvol (mGy)']>10]
        else:
            pass
        
        if name == 'Head':
            for row in df.index:
                if ('PERFUSION' in str(df.loc[row, 'Description'])) | ('ANGIO' in str(df.loc[row, 'Description'])):
                    df = df.drop(row)
                else:
                    pass
            df = df.drop('Description', axis = 1)
        df.sort_values(by= 'Date')
        for row in df.index:
            df.loc[row, 'Month'] = int(df.loc[row, 'Date'].month)
        for row in df.index:
            df.loc[row, 'Year'] = int(df.loc[row, 'Date'].year)
        for row in df.index:
            if df.loc[row, 'Month'] <10:
                df.loc[row, 'year-month'] = str(int(df.loc[row, 'Year'])) + '-0' + str(int(df.loc[row,'Month']))
            else:
                df.loc[row, 'year-month'] = str(int(df.loc[row, 'Year'])) + '-' + str(int(df.loc[row,'Month']))
        df.drop('Date', axis = 1, inplace = True)
        # df.to_excel(r'Z:\Emi\Imalogix\AirKermaTEST.xlsx')
        # df2 = df2.groupby('Month').agg({'CTDIvol (mGy)':['median', 'max']})
        df2 = df.groupby(['year-month']).agg({'CTDIvol (mGy)':['median', 'max']})
        
        # current = list(set(df['year-month']))
        # current = [int(i) for i in current]
        # new = [calendar.month_name[i] for i in current]
        
        # current = [i for i in range(1, len(new)+1)]
        # current = list(set(oglist_collect))
        # new = list(set(df['year-month']))
        boxplot =df.boxplot(column = 'CTDIvol (mGy)', by = 'year-month', \
                            grid = False, figsize = (10,4), medianprops={'linewidth': 3, 'color': 'red'}, \
                            showfliers=False, return_type = 'dict')
        
        # plt.xticks(current, current, rotation = 45)   
        plt.xticks(rotation=45)

        bottom, top = plt.ylim()
        plt.ylim(0, top)   

        plt.ylabel('CTDIvol (mGy)')
        plt.xlabel('Month')
        
        
        plt.title('CTDIvol by Month')
        plt.suptitle('')
        plt.savefig(bpfilename, bbox_inches = "tight")
        plt.close()
        
        
        
        f, ax = plt.subplots(figsize=(10,3))
        maxplot = plt.bar(df2.index.tolist(), df2[('CTDIvol (mGy)', 'max')])
        plt.xticks(rotation=45)
        # current = list(set(df['year-month']))
        
        # plt.xticks(current, current, rotation = 45)   
        plt.title('Max CTDIvol by Month')
        plt.ylabel('CTDIvol (mGy)')
        plt.xlabel('Month')
    
        
        plt.savefig(chartfilename, bbox_inches = "tight")

        df2.to_excel(writer, sheet_name = name)
        ws = writer.sheets[name]
        ws.insert_image('E1', '', {'image_data': bpfilename})
        ws.insert_image('E25', '', {'image_data': chartfilename})
        plt.close()
    summary_df.to_excel(writer, sheet_name = 'Summary', index = False)
writer.save
xl.close()
        
