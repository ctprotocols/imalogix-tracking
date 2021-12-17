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

path = r'Z:\Emi\Code\Imalogix_tracking_code\Air Kerma\AirKerma_Tracking_Enter_Data.xlsx'

# path = r'Z:\Emi\Code\Imalogix_tracking_code\CATH\CATH_Tracking_Enter_Data.xlsx'
plotpath = r'Z:\Emi\Code\Imalogix_tracking_code\CATH\CATH_Tracking_Output.xlsx'


xl = pd.ExcelFile(path)
name = 'CATH'
df = pd.read_excel(path, sheet_name = name)

scanner_names = list(set(df['Scanner'].tolist()))
scanner_names.sort()
with pd.ExcelWriter(plotpath, engine = 'xlsxwriter') as writer:
    for scannername in scanner_names:
        bpfilename = io.BytesIO()
        chartfilename = io.BytesIO()
        newdf = df.loc[df['Scanner'] == scannername]
        newdf = newdf[['Air Kerma (mGy)', 'Date']]
        newdf.sort_values(by= 'Date')
        for row in newdf.index:
            newdf.loc[row, 'Month'] = int(newdf.loc[row, 'Date'].month)
        for row in newdf.index:
            newdf.loc[row, 'Year'] = int(newdf.loc[row, 'Date'].year)
        for row in newdf.index:
            if newdf.loc[row, 'Month'] <10:
                newdf.loc[row, 'year-month'] = str(int(newdf.loc[row, 'Year'])) + '-0' + str(int(newdf.loc[row,'Month']))
            else:
                newdf.loc[row, 'year-month'] = str(int(newdf.loc[row, 'Year'])) + '-' + str(int(newdf.loc[row,'Month']))
        
        
        newdf.drop('Date', axis = 1, inplace = True)
        # newdf.to_excel(r'Z:\Emi\Imalogix\AirKermaTEST.xlsx')
        
        df2 = newdf.groupby('year-month').agg({'Air Kerma (mGy)':['median', 'max']})

        # current = list(set(newdf['Month']))
        # current = [int(i) for i in current]
        # new = [calendar.month_name[i] for i in current]
        
        # current = [i for i in range(1, len(new)+1)]
        
        allmonths = list(set(newdf['year-month']))
        allmonths.sort()
        monthlists = []
        for month in allmonths:
            monthlists.append(newdf.loc[newdf['year-month']==month]['Air Kerma (mGy)'].tolist())
            
            
            
        fig, ax = plt.subplots()
        ax.boxplot(monthlists, showfliers=False, boxprops={'color':'tomato'}, capprops={'color':'tomato'},whis = 0, medianprops={'linewidth': 2, 'color': 'red'})
        ax.tick_params(axis='y', labelcolor='red')

        plt.xticks(plt.xticks()[0], allmonths)
        plt.xticks(rotation = 60)   
        bottom, top = plt.ylim()
        plt.ylim(0, top)   

        plt.ylabel('Air Kerma (mGy)', color = 'red')
        plt.xlabel('Month')
        
        twin1 = ax.twinx()
        xlabels = [i for i in range(1, len(df2.index.tolist())+1)]
        twin1.scatter(xlabels, df2[('Air Kerma (mGy)', 'max')], alpha = 0.8, color = 'blue')
        twin1.tick_params(axis='y', labelcolor='blue')
        twin1.set_ylabel('Max Air Kerma (mGy)', color = 'blue')
        # plt.ylim(0, top)   

        plt.title('Air Kerma by Month')

        # plt.show()  
            
            
            
        # plt.suptitle('')
        plt.savefig(bpfilename, bbox_inches = "tight", dpi = 350)
        # plt.close()
        
        
    
        df2.to_excel(writer, sheet_name = scannername)
        ws = writer.sheets[scannername]
        ws.insert_image('E1', '', {'image_data': bpfilename})
        
writer.save
