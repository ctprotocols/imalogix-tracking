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
import seaborn as sns

path = r'Z:\Emi\Code\Imalogix_tracking_code\IR\IR_Tracking_Enter_Data.xlsx'
plotpath = r'Z:\Emi\Code\Imalogix_tracking_code\IR\IR_Tracking_Output.xlsx'


xl = pd.ExcelFile(path)
sheet_names = xl.sheet_names

df_all = pd.DataFrame()
with pd.ExcelWriter(plotpath, engine = 'xlsxwriter') as writer:
    for name in sheet_names:
        df = pd.read_excel(path, sheet_name = name)
        df = df[['Exam Protocol','Air Kerma (mGy)', 'Date']]
        df.insert(0, 'Location', name)
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
        df_all = pd.concat([df_all, df])

        # df.to_excel(r'Z:\Emi\Imalogix\AirKermaTEST.xlsx')
        
    VAS = df_all.loc[(df_all['Location'] == 'Vas1') | (df_all['Location'] == 'Vas2')]
    NEURO = df_all.loc[(df_all['Location'] == 'Neuro1') | (df_all['Location'] == 'Neuro2')]    
    MYELO = df_all.loc[df_all['Location'] == 'Myelo']
    AHSP =  df_all.loc[df_all['Location'] == 'AHSP']   
    
    df_list = [VAS, NEURO, MYELO, AHSP ]
    df_name_list = ['VAS', 'NEURO', 'MYELO', 'AHSP' ]

    for idx1, j in enumerate(df_list):
        linecount = 0

        df2 = j.groupby(['year-month', 'Exam Protocol', 'Location']).agg({'Air Kerma (mGy)':['median', 'max']})
        df2.to_excel(writer, sheet_name = df_name_list[idx1])

        protocols = list(set(j['Exam Protocol']))
        for idx2, p in enumerate(protocols):
            bpfilename = io.BytesIO()

            df_p = j.loc[j['Exam Protocol'] == p]
            months = sorted(list(set(df_p['year-month'])))
            
            graph_df = pd.DataFrame()
            for m in months:
                df_p_bymonth = df_p.loc[df_p['year-month'] == m]
                if df_name_list[idx1] == 'VAS':
                    vas1 = df_p_bymonth.loc[df_p_bymonth['Location'] == 'Vas1']
                    vas2 = df_p_bymonth.loc[df_p_bymonth['Location'] == 'Vas2']
                    if vas1.shape[0] < 3:
                        pass
                    else:
                        graph_df = pd.concat([graph_df, vas1])
                    if vas2.shape[0] < 3:
                        pass
                    else:
                        graph_df = pd.concat([graph_df, vas2])
                elif df_name_list[idx1] == 'NEURO':
                    neu1 = df_p_bymonth.loc[df_p_bymonth['Location'] == 'Neuro1']
                    neu2 = df_p_bymonth.loc[df_p_bymonth['Location'] == 'Neuro2']
                    if neu1.shape[0] < 3:
                        pass
                    else:
                        graph_df = pd.concat([graph_df, neu1])
                    if neu2.shape[0] < 3:
                        pass
                    else:
                        graph_df = pd.concat([graph_df, neu2])
                else:
                    if df_p_bymonth.shape[0] < 3:
                        pass
                    else:
                        graph_df = pd.concat([graph_df, df_p_bymonth])
            print(graph_df)
            if graph_df.empty:
                continue
            
            fig, ax = plt.subplots() 
            sns.boxplot(y='Air Kerma (mGy)', x='year-month', 
                     data=graph_df, 
                     palette="colorblind",
                     hue='Location', showfliers=False)
            plt.xticks(rotation = 45)
            plt.title(p)
            plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.)
            plt.savefig(bpfilename, bbox_inches = "tight")
            plt.close()
            

            ws = writer.sheets[df_name_list[idx1]]
            
            cellname = 'G' + str(5 + linecount*20)
            ws.insert_image(cellname, '', {'image_data': bpfilename})
            linecount += 1
writer.save
xl.close()
        
        # df2 = df.groupby('Month').agg({'Air Kerma (mGy)':['median', 'max']})

        # current = list(set(df['Month']))
        # current = [int(i) for i in current]
        # new = [calendar.month_name[i] for i in current]
        
        # current = [i for i in range(1, len(new)+1)]
        
        # dict_df = {}
        # data_dict = []
        # for p in protocols:
        #     df_p = df.loc[df['Exam Protocol'] == p]
        #     dict_df[p] = [df_p['Month'].tolist(), df_p['Air Kerma (mGy)'].tolist()]
        #     data_dict.append(df_p['Air Kerma (mGy)'].tolist())
            # df_p.boxplot(column = 'Air Kerma (mGy)', by = 'Month', \
            #                         grid = False, figsize = (10,4), medianprops={'linewidth': 3, 'color': 'red'}, \
            #                         showfliers=False, return_type = 'dict')

        # fig = figure()
        # ax = axes()
        # hold(True)
        
        # bp = boxplot(A, positions = [1, 2], widths = 0.6)

        # boxplot =df.boxplot(column = 'Air Kerma (mGy)', by = 'Month', \
        #                     grid = False, figsize = (10,4), medianprops={'linewidth': 3, 'color': 'red'}, \
        #                     showfliers=False, return_type = 'dict')
        
        # plt.xticks(current, new, rotation = 45)   
        # bottom, top = plt.ylim()
        # plt.ylim(0, top)   

        # plt.ylabel('Air Kerma (mGy)')
        # plt.xlabel('Month (2020)')
        
        
        # plt.title('Air Kerma by Month')
        # plt.suptitle('')
        # plt.savefig(bpfilename, bbox_inches = "tight")
        # plt.close()
        
        
        
#         f, ax = plt.subplots(figsize=(10,3))
#         maxplot = plt.bar(df2.index.tolist(), df2[('Air Kerma (mGy)', 'max')])
#         current = list(set(df['Month']))
        
#         plt.xticks(current, new, rotation = 45)   
#         plt.title('Max Air Kerma by Month')
#         plt.ylabel('Air Kerma (mGy)')
#         plt.xlabel('Month (2020)')
    
        
#         plt.savefig(chartfilename, bbox_inches = "tight")

#         df2.to_excel(writer, sheet_name = name)
#         ws = writer.sheets[name]
#         ws.insert_image('E1', '', {'image_data': bpfilename})
#         ws.insert_image('E25', '', {'image_data': chartfilename})
        
        
        
