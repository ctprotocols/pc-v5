# -*- coding: utf-8 -*-
"""
Created on Fri Dec  4 07:03:11 2020

@author: EastmanE
"""

import win32com.client as win32

import pandas as pd
import xlsxwriter
import numpy as np
import os
import datetime as dt


def GE(oldfilename, newfilename):
    path_results = os.path.join(os.path.split(newfilename)[0], 'Comparison.xlsx')
    
    df_old = pd.read_excel(oldfilename)
    df_new = pd.read_excel(newfilename)
    
    
    final_df = pd.DataFrame()
    
    #load all protocol names into a list
    protocols_old = list(set(df_old['Protocol'].tolist()))
    protocols_new = list(set(df_new['Protocol'].tolist()))
    
    #find protocols in both files
    same_prot = [i for i in protocols_old if i in protocols_new]
    #find protocols in old file that are not in new file (removed)
    removed = [i for i in protocols_old if i not in protocols_new]
    #find protocols in new file that are not in old file (added)
    added = [i for i in protocols_new if i not in protocols_old]
    
    # =============================================================================
    # Find protocols that have been renumbered
    # =============================================================================
    
    added_edit = {}
    for addidx, i in enumerate(added):
        i_list = i.split()
        del i_list[2]
        for idx, ii in enumerate(i_list):
            ii = ii.replace('*', '')
            ii = ii.replace('-', '')
            i_list[idx] = ii
            if ii == '':
                del i_list[idx]
        added_edit[addidx] = i_list
        
    removed_edit = {}
    for remidx, i in enumerate(removed):
        i_list = i.split()
        del i_list[2]
        for idx, ii in enumerate(i_list):
            ii = ii.replace('*', '')
            ii = ii.replace('-', '')
            i_list[idx] = ii
            if ii == '':
                del i_list[idx]
        removed_edit[remidx] = i_list
    
    
    
    same_edit = {}
    remove_add_dict = []
    remove_rem_dict = []
    for i in added_edit:
        for j in removed_edit:
            if added_edit[i] == removed_edit[j]:
                same_edit[i] = j
                remove_add_dict.append(i)
                remove_rem_dict.append(j)
    
    
    
    for i in remove_add_dict:
        del added_edit[i]
    for i in remove_rem_dict:
        del removed_edit[i]
    
    #Find approx matches
    
    remove_add_dict = []
    remove_rem_dict = []
    for i in added_edit:
        w_add = added_edit[i].copy()
        for j in removed_edit:
            w_add_2 = added_edit[i].copy()
            w_rem = removed_edit[j].copy()
            for k in w_add:
                for l in w_rem:
                    if k == l:
                        w_rem.remove(l)
                        w_add_2.remove(k)
                        break
                    else:
                        pass
            if (len(w_add) > 6) | (len(w_rem) > 6):
                if (len(w_rem) <3) & (len(w_add_2) <3):
                    same_edit[i] = j
                    remove_add_dict.append(i)
                    remove_rem_dict.append(j)
    
                    break
            else:
                if (len(w_rem) <2) & (len(w_add_2) <2):
                    same_edit[i] = j
                    remove_add_dict.append(i)
                    remove_rem_dict.append(j)
                    break        
    
    for i in remove_add_dict:
        del added_edit[i]
    for i in remove_rem_dict:
        del removed_edit[i]
    
    
    
    
    same_edit_names = {}
    
    for i in same_edit:
        newname = str(np.random.randint(10000000, 99999999)) 
        #make sure it is not a duplicate number
        while newname in same_edit_names.keys():
            newname = str(np.random.randint(10000000, 99999999))
        same_edit_names[newname] = [removed[same_edit[i]], added[i]]
        df_old.replace(removed[same_edit[i]], newname, inplace = True)
        df_new.replace(added[i], newname, inplace = True)
    
    
    #Remove these scans from the removed and added list since they don't belong
    delfromrem= [removed[same_edit[i]] for i in same_edit]
    delfromadd = [added[i] for i in same_edit]
    removed = [i for i in removed if i not in delfromrem]
    added = [i for i in added if i not in delfromadd]
    
    
            
    #load all protocol names into a list
    protocols_old = list(set(df_old['Protocol'].tolist()))
    protocols_new = list(set(df_new['Protocol'].tolist()))
    
    #find protocols in both files
    same_prot = [i for i in protocols_old if i in protocols_new]
    
    
    #scans removed to put at the end of the document
    deleted_df = pd.DataFrame()
    for p in removed:
        deleted_df = pd.concat([deleted_df, df_old.loc[df_old['Protocol'] == p]], ignore_index = True)
    
    
    removed_df = pd.DataFrame()
    added_df = pd.DataFrame()
    changes_df = pd.DataFrame(columns=['Protocol', 'Scan/Recon Type', 'Parameter', 'Old', 'New'])
    
    # =============================================================================
    # Find protocols that have been changed
    # =============================================================================
    row_of_change = []
    changes = []
    changes_row = 0
    for protocol in same_prot:
        rest_19 = pd.DataFrame()
        rest_20 = pd.DataFrame()
    
        #Isolate the protocol- last year
        p_last = df_old.loc[df_old['Protocol'] == protocol].copy()
        p_last.reset_index(drop = True, inplace = True)
        #Fill NaN with 'NA'. Otherwise, df.equals will fail
        p_last.fillna('NA', inplace = True)
        #Iolate the protocol- this year
        p_new = df_new.loc[df_new['Protocol'] == protocol].copy()
        p_new.reset_index(drop = True, inplace = True)
        p_new.fillna('NA', inplace = True)
        
        # if not p_new.equals(p_last):
        #     changes.append(protocol)
        # else:
        #     pass
        
        #Find and get rid of rows that are the same. 
        for r19 in p_last.index:
            #Look at the prior year, one line at a time
            row_19 = p_last.loc[r19, :].copy()
            #Compare the line from the prior year to each line in this year's
            #protocol until it finds a match
            for r20 in p_new.index:
                row_20 = p_new.loc[r20, :].copy()
                #If it finds a match, drop the row from both of the datafranes.
                #Whatever is leftover in the dataframe is a new change.
                if row_19.equals(row_20):
                    p_last.drop(r19, axis=0, inplace = True)
                    p_new.drop(r20, axis=0, inplace = True)
                    break
        #If the dataframes from the previous and current year are empty, a match was
        #found for every line. The protocol stays the same
        if (p_last.empty) & (p_new.empty):  
            continue
        #If last year's is empty but this year's is not, something new was added.
        elif (p_last.empty) & (not p_new.empty):
            added_df = pd.concat([added_df, p_new], axis = 0)
        #if this year's is empty but last year's is not, something was removed from
        #the old protocol.
        elif (p_new.empty) & (not p_last.empty):
            removed_df = pd.concat([removed_df, p_last], axis = 0)
        #If they both have lines in them, there could be additions, deletions, or edits
        #We need to categorize them through further analysis.
        else:
            rest_19 = pd.concat([rest_19, p_last])
            rest_20 = pd.concat([rest_20, p_new])
    
        #Sort through whatever is leftover
        if rest_19.empty and rest_20.empty:
            continue
        else:
            for row_m in rest_19.index:
                for row_n in rest_20.index:
                    if rest_19.loc[row_m, 'Scan/Recon Type'] == rest_20.loc[row_n, 'Scan/Recon Type']:
                        for col in rest_19.columns:
                            if rest_19.loc[row_m, col] != rest_20.loc[row_n, col]:
                                #Ignore differences within 5%
                                if (type(rest_19.loc[row_m, col]) == float) & (type(rest_20.loc[row_n, col]) == float):
                                    if abs(rest_19.loc[row_m, col] - rest_20.loc[row_n, col]) < 0.75:
                                        continue
                                    else:
                                        pass
                                #ignore change in naming convention (e.g. 25 ->25.0 D)
                                elif (str(rest_19.loc[row_m, col]) in str(rest_20.loc[row_n, col])) | (str(rest_20.loc[row_n, col]) in str(rest_19.loc[row_m, col])):
                                    continue
                                else:
                                    pass
                                #Collect row of the change within the protocol to use with iloc later
                                row_of_change.append(row_m)
                                row_of_change.append(row_n)
                                
                                #Collect summary of changes in changes_df
                                changes_df.loc[changes_row, 'Protocol'] = rest_19.loc[row_m, 'Protocol']
                                changes_df.loc[changes_row,'Scan/Recon Type'] = rest_19.loc[row_m, 'Scan/Recon Type']
                                changes_df.loc[changes_row, 'Parameter'] = col
                                changes_df.loc[changes_row, 'Old'] = rest_19.loc[row_m, col]
                                changes_df.loc[changes_row, 'New'] = rest_20.loc[row_n, col]
                                changes_row +=1
                            else:
                                pass
                        #drop the row after we analyze it. 
                        rest_19.drop(row_m, axis = 0, inplace = True)
                        rest_20.drop(row_n, axis = 0, inplace = True)
                        break
                    else:
                        pass
            #If there are any rows left, these are the rows that were added or removed.
            if not rest_19.empty:
                removed_df = pd.concat([removed_df, rest_19])
            if not rest_20.empty:
                added_df = pd.concat([added_df, rest_20])
    
    changed_protocols = list(set(changes_df['Protocol']))
    #Add in names of protocols with whole scans removed
    if not added_df.empty:
        addnames = list(set(added_df['Protocol']))
        changed_protocols.extend(addnames)

    if not removed_df.empty:
        removenames = list(set(removed_df['Protocol']))
        changed_protocols.extend(removenames)
    
    # =============================================================================
    # Collect locations requiring formatting for individual changes
    # =============================================================================
    
    unchanged_df = pd.DataFrame()
    for p in protocols_new:
        if p in changed_protocols:
            original_df = df_old.loc[df_old['Protocol'] == p].copy()
            oldname = p + ' (Old)'
            original_df['Protocol'] = oldname
            
            new_df = df_new.loc[df_new['Protocol'] == p].copy()
            newname = p + ' (New)'
            new_df['Protocol'] = newname
    
            final_df = pd.concat([final_df, original_df], ignore_index=True)
            final_df = pd.concat([final_df, new_df], ignore_index = True)
                    
        if p in added:
            final_df = pd.concat([final_df, df_new.loc[df_new['Protocol'] == p]], ignore_index=True)
        else:
            unchanged_df = pd.concat([unchanged_df, df_new.loc[df_new['Protocol'] == p]], ignore_index=True)
    
    remove_location = final_df.tail(n=1).index.tolist()[0]
    final_df = pd.concat([final_df, deleted_df], ignore_index = True)
    
    
    # =============================================================================
    # Format protocols
    # =============================================================================
    
    #drop blank columns
    final_df.dropna(how = 'all', axis = 1, inplace = True)
    #collect index values of the first instance of a protocol name
    insertloc = []
    for row in final_df.index[:-1]:
        if final_df.loc[row, 'Protocol'] != final_df.loc[row+1, 'Protocol']:
            insertloc.append(row)
        else:
            pass
    
    ###REPEAT FOR UNCHANGED###
    #drop blank columns
    unchanged_df.dropna(how = 'all', axis = 1, inplace = True)
    #collect index values of the first instance of a protocol name
    insertloc_nc = []
    for row in unchanged_df.index[:-1]:
        if unchanged_df.loc[row, 'Protocol'] != unchanged_df.loc[row+1, 'Protocol']:
            insertloc_nc.append(row)
        else:
            pass
    
    
    addvalue = 0
    formatloc = []
    
    for i, idx in enumerate(insertloc):
        #update the index value with the add value
        idx = idx + addvalue
        formatloc.append(idx+1)
        #addvalue takes into account the additional rows we will add
        addvalue = addvalue + 1
        #Create a blank row to insert between protocols at the locations found in formatloc
        blankrow = pd.DataFrame([[''] * len(final_df.columns)], columns = final_df.columns)
        #Place the protocol name in the 2nd column of the blank row. We will drop the first column of protocol names later
        blankrow.iloc[0,1] = final_df.iloc[idx+1, 0]
        #insert the blank row
        final_df = pd.concat([final_df[:idx+1], blankrow, final_df[idx+1:]], ignore_index = False)
    
    ###REPEAT NO CHANGES####
    addvalue_nc = 0
    formatloc_nc = []
    
    for i, idx in enumerate(insertloc_nc):
        #update the index value with the add value
        idx = idx + addvalue_nc
        formatloc_nc.append(idx+1)
        #addvalue takes into account the additional rows we will add
        addvalue_nc = addvalue_nc + 1
        #Create a blank row to insert between protocols at the locations found in formatloc
        blankrow_nc = pd.DataFrame([[''] * len(unchanged_df.columns)], columns = unchanged_df.columns)
        #Place the protocol name in the 2nd column of the blank row. We will drop the first column of protocol names later
        blankrow_nc.iloc[0,1] = unchanged_df.iloc[idx+1, 0]
        #insert the blank row
        unchanged_df = pd.concat([unchanged_df[:idx+1], blankrow_nc, unchanged_df[idx+1:]], ignore_index = False)
    
    
        
    #First protocol name
    blankrow = pd.DataFrame([[''] * len(final_df.columns)], columns = final_df.columns)
    blankrow.iloc[0,1] = final_df.iloc[0,0]
    final_df = pd.concat([blankrow, final_df], ignore_index = True)
    
    
    ###REPEAT FOR NO CHANGES###
    #First protocol name
    blankrow_nc = pd.DataFrame([[''] * len(unchanged_df.columns)], columns = unchanged_df.columns)
    blankrow_nc.iloc[0,1] = unchanged_df.iloc[0,0]
    unchanged_df = pd.concat([blankrow_nc, unchanged_df], ignore_index = True)
    
    
    #Collect the row and column values for each changed parameter
    rowlist = []
    collist = []
        
    for row in changes_df.index:
        oldrow = final_df.loc[final_df['Protocol'] == changes_df.loc[row, 'Protocol']+ ' (Old)']
        oldrow = oldrow.iloc[row_of_change[2*row], :]
        column_name = changes_df.loc[row, 'Parameter']
        oldrow = oldrow.name
        rowlist.append(oldrow)
        oldcol = final_df.columns.get_loc(column_name)
        collist.append(oldcol)
    
        
        newrow = final_df.loc[final_df['Protocol'] == changes_df.loc[row, 'Protocol'] + ' (New)']
        newrow = newrow.iloc[row_of_change[2*row+1], :]
        
        newrow = newrow.name
        rowlist.append(newrow)
        newcol = final_df.columns.get_loc(column_name)
        collist.append(newcol)
        
    
    
    # =============================================================================
    # Collect locations requiring formatting for entire rows (new scan/recons added)
    # =============================================================================
    removed_df.reset_index(inplace = True)
    delete_whole_row = []
    for row in removed_df.index:
        removed_scan = final_df.loc[final_df['Protocol'] == removed_df.loc[row, 'Protocol']+ ' (Old)']
        removed_scan = removed_scan.iloc[removed_df.loc[row, 'index'], :]
        delete_whole_row.append(removed_scan.name)
    
    added_df.reset_index(inplace = True)
    add_whole_row = []
    for row in added_df.index:
        added_scan = final_df.loc[final_df['Protocol'] == added_df.loc[row, 'Protocol']+ ' (New)']
        added_scan = added_scan.iloc[added_df.loc[row, 'index'], :]
        add_whole_row.append(added_scan.name)
    
    
    
    final_df.drop('Protocol', axis = 1, inplace = True)
    unchanged_df.drop('Protocol', axis = 1, inplace = True)
    
    
    for i in same_edit_names:
        final_df.replace(i+ ' (Old)', same_edit_names[i][0] + ' (Old)', inplace = True)
        final_df.replace(i+ ' (New)', same_edit_names[i][1] + ' (New- Renamed)', inplace = True)
    
    for i in same_edit_names:
        unchanged_df.replace(i, same_edit_names[i][1] + (' (Renamed)'), inplace = True)
    
    formatloc = [i+2 for i in formatloc]
    formatloc.insert(0,1)
    
    protocolnames = [final_df.iloc[index-1, 0] for index in formatloc]
    
    
    #REPEAT FOR NO CHANGES#
    formatloc_nc = [i+2 for i in formatloc_nc]
    formatloc_nc.insert(0,1)
    
    protocolnames_nc = [unchanged_df.iloc[index-1, 0] for index in formatloc_nc]
    
    
    
    sheetname = 'Changes'
    sheetname_unchanged = 'No Changes'
    with pd.ExcelWriter(path_results, engine = 'xlsxwriter') as writer:
        
        final_df.to_excel(writer, sheet_name = sheetname, index = False, engine = 'openpyxl')
        workbook = writer.book
        worksheet = writer.sheets[sheetname]
        
        
        format_title_normal = workbook.add_format({'bold': 1,'border': 1,'align': 'left','valign': 'vcenter','fg_color': 'gray', 'font_color':'white'})
        format_plain = workbook.add_format({'border': 1,'align': 'left','valign': 'vcenter', 'font_color':'black'})
    
        format_title_new = workbook.add_format({'bold': 1,'border': 1,'align': 'left','valign': 'vcenter','fg_color': 'green', 'font_color':'white'})
        format_edits = workbook.add_format({'bold': 1,'bg_color': 'yellow'})
        format_new = workbook.add_format({'border': 1,'align': 'left','valign': 'vcenter', 'bg_color': 'green', 'font_color':'white'})
        
        format_title_removed = workbook.add_format({'bold': 1,'border': 1,'align': 'left','valign': 'vcenter','fg_color': 'gray', 'font_color':'black'})
        format_removed = workbook.add_format({'border': 1,'align': 'left','valign': 'vcenter','bg_color': 'gray', 'font_color':'black'})
        format_delete_row = workbook.add_format({'border': 1,'align': 'left','valign': 'vcenter','bg_color': 'gray', 'font_color':'red'})
    
        
        # worksheet.set_column(len(final_df.columns),30, 5, format_plain)
    
        for nameidx, idx in enumerate(formatloc):
            protname = protocolnames[nameidx]
    
            if nameidx < len(formatloc)-1:
                if protname in added:
                    cell_format = format_title_new
                elif protname in removed:
                    cell_format = format_title_removed
                else:
                    cell_format = format_title_normal
        
                worksheet.merge_range(idx, 0, idx, final_df.shape[1]-1, protname, cell_format)
                
                if protname in added:
                    worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, formatloc[nameidx+1]-1, len(final_df.columns)-1), {'type': 'no_errors', 'format': format_new})
                elif protname in removed:
                    worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, formatloc[nameidx+1]-1, len(final_df.columns)-1), {'type': 'no_errors', 'format': format_removed})
                
                else:
                    worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, formatloc[nameidx+1]-1, len(final_df.columns)-1), {'type': 'no_errors', 'format': format_plain})
            else:
                if protname in added:
                    cell_format = format_title_new
                elif protname in removed:
                    cell_format = format_title_removed
                else:
                    cell_format = format_title_normal
        
                worksheet.merge_range(idx, 0, idx, final_df.shape[1]-1, protname, cell_format)
                
                if protname in added:
                    worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, final_df.shape[0], len(final_df.columns)-1), {'type': 'no_errors', 'format': format_new})
                elif protname in removed:
                    worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, final_df.shape[0], len(final_df.columns)-1), {'type': 'no_errors', 'format': format_removed})
                
                else:
                    worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, final_df.shape[0], len(final_df.columns)-1), {'type': 'no_errors', 'format': format_plain})
    
                
        for i, (r, c) in enumerate(zip(rowlist, collist)): 
            worksheet.conditional_format(xlsxwriter.utility.xl_range(r+1, c-1, r+1, c-1), {'type': 'no_errors', 'format': format_edits})
            
        for y1 in delete_whole_row:
            worksheet.conditional_format(xlsxwriter.utility.xl_range(y1+1, 0, y1+1, len(final_df.columns)-1), {'type': 'no_errors', 'format': format_delete_row})
        for y2 in add_whole_row:
            worksheet.conditional_format(xlsxwriter.utility.xl_range(y2+1, 0, y2+1, len(final_df.columns)-1), {'type': 'no_errors', 'format': format_edits})
            
        unchanged_df.to_excel(writer, sheet_name = sheetname_unchanged, index = False, engine = 'openpyxl')
        workbook = writer.book
        worksheet = writer.sheets[sheetname_unchanged]
        for nameidx, idx in enumerate(formatloc_nc):
            protname = protocolnames_nc[nameidx]
            if nameidx < len(formatloc_nc)-1:
                cell_format = format_title_normal
                worksheet.merge_range(idx, 0, idx, unchanged_df.shape[1]-1, protname, cell_format)
                worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, formatloc_nc[nameidx+1]-1, len(unchanged_df.columns)-1), {'type': 'no_errors', 'format': format_plain})
            else:
                cell_format = format_title_normal
                worksheet.merge_range(idx, 0, idx, unchanged_df.shape[1]-1, protname, cell_format)
                worksheet.conditional_format(xlsxwriter.utility.xl_range(idx+1, 0, unchanged_df.shape[0], len(unchanged_df.columns)-1), {'type': 'no_errors', 'format': format_plain})
    
    
    exl = win32.Dispatch('Excel.Application')
    wb = exl.Workbooks.Open(Filename = path_results)
    ws = wb.Worksheets(sheetname)
    ws1 = wb.Worksheets(sheetname_unchanged)
    ws1.Columns('A:Z').AutoFit()
    ws.Columns('A:Z').AutoFit()
    wb.Save()
    wb.Close()
    exl.Quit()        





