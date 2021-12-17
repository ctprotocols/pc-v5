# -*- coding: utf-8 -*-
"""
Created on Thu Oct  1 15:22:04 2020

@author: EastmanE

"""
import pandas as pd

# filepath = r'Z:\Emi\Prots\Toshformat.xlsx'
# despath = r'Z:\Emi\Prots\Toshform.xlsx'

def PDF_formatter(filepath, despath):
    import win32com.client as win32
    import xlsxwriter
    with pd.ExcelWriter(despath, engine = 'xlsxwriter') as writer:
        machinenames = pd.ExcelFile(filepath).sheet_names
        # formatlocs = []
        for name in machinenames:
            df = pd.read_excel(filepath, sheet_name = name)
            df.dropna(how='all', axis=1, inplace = True)
            insertloc =[]
            for row in df.index[:-1]: 
                if df.loc[row, 'Protocol'] != df.loc[row + 1, 'Protocol']:
                    insertloc.append(row)
                else:
                    pass
            # insertloc.insert(0, 0)
            
            addvalue = 0
            formatloc = []
            protocolnames = []
            for i, idx in enumerate(insertloc):
                idx = idx + addvalue
                formatloc.append(idx+1)
                addvalue = addvalue + 1
                blankrow = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)       #first protocol name
                blankrow.iloc[0, 1] = df.iloc[idx+1, 0]
                df = pd.concat([df[:idx+1], blankrow, df[idx+1:]], ignore_index= False)
            
            blankrow = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)       #first protocol name
            blankrow.iloc[0, 1] = df.iloc[0, 0]
            df = pd.concat([blankrow, df], ignore_index= False)
    
            df.drop('Protocol', axis = 1, inplace = True)
            df.reset_index(inplace = True)
            df.drop('index', axis = 1, inplace = True)
            
            formatloc = [i+2 for i in formatloc]
            formatloc.insert(0, 1)
            
            for index in formatloc:
                protocolnames.append(df.iloc[index-1, 0])
                
            df.to_excel(writer, sheet_name = name, index = False)
            workbook = writer.book
            worksheet = writer.sheets[name]
            border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
            worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df), len(df.columns)-1), {'type': 'no_errors', 'format': border_fmt})
            cell_format = workbook.add_format({'bold': 1,'border': 1,'align': 'left','valign': 'vcenter','fg_color': 'gray', 'font_color':'white'})
            for formidx, k in enumerate(formatloc):
                protname = protocolnames[formidx]
                worksheet.merge_range(k, 0, k, df.shape[1]-1, protname, cell_format)
            print(name)
        
    xl = win32.Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(Filename = despath)
    for name in machinenames:
        ws = wb.Worksheets(name)
        ws.Columns('A:Z').AutoFit()
    wb.Save()
    wb.Close()
    xl.Quit()        