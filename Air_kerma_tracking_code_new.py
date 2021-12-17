# -*- coding: utf-8 -*-
"""
Created on Fri Oct 30 13:25:11 2020

@author: EastmanE
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import calendar
import io 

# matplotlib.rcParams['figure.dpi'] = 350 

path = r'Z:\Emi\Code\Imalogix_tracking_code\Air Kerma\AirKerma_Tracking_Enter_Data.xlsx'
plotpath = r'Z:\Emi\Code\Imalogix_tracking_code\Air Kerma\AirKerma_Tracking_Output.xlsx'


xl = pd.ExcelFile(path)
sheet_names = xl.sheet_names
with pd.ExcelWriter(plotpath, engine = 'xlsxwriter') as writer:
    for name in sheet_names:
        bpfilename = io.BytesIO()
        # chartfilename = io.BytesIO()
        
        df = pd.read_excel(path, sheet_name = name)
        df = df[['Air Kerma (mGy)', 'Date']]
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
        
        df2 = df.groupby('year-month').agg({'Air Kerma (mGy)':['median', 'max']})

        # current = list(set(df['Month']))
        # current = [int(i) for i in current]
        # new = [calendar.month_name[i] for i in current]
        
        # current = [i for i in range(1, len(new)+1)]
    
    
        allmonths = list(set(df['year-month']))
        allmonths.sort()
        monthlists = []
        for month in allmonths:
            monthlists.append(df.loc[df['year-month']==month]['Air Kerma (mGy)'].tolist())
            
            
            
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
        
        
    
        df2.to_excel(writer, sheet_name = name)
        ws = writer.sheets[name]
        ws.insert_image('E1', '', {'image_data': bpfilename})
        
writer.save
        
        
