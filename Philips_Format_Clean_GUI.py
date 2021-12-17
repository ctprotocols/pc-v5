# =============================================================================
# Files must be filed: 'Main Folder'\'Machine name containing the vendor name (GE, Toshiba, or Siemens)\2020\file in xlsx format
# =============================================================================
# -*- coding: utf-8 -*-
"""
Created on Fri Sep  4 07:30:33 2020

@author: EastmanE
"""
import os
import datetime as dt
import pandas as pd
import openpyxl as xl
import numpy as np
from Philips_settings import finalheaders, keyword1, keyword2, originalheaders, foldpath, finalpath, despath
import PDFformatter

def Philips(filename, machinename, typeofscanner):

    year = dt.datetime.today().strftime('%Y')
    # year = '2019'
    
    dirname = os.path.dirname(filename)
    finalfilename = 'Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    
    finalpath = os.path.join(dirname, finalfilename)
    
    year = dt.datetime.today().strftime('%Y')
    
    exportfile = filename
    
    # =============================================================================
    #Define Functions
    # =============================================================================
    
    # =============================================================================
    # clean_function removes blank and duplicate columns from the df to allow concatenation
    # =============================================================================
    
    def clean_function(dfname):
        blankspace = [col for col in dfname.columns if col == ' ']  #This list will be empty unless there is a column named ' '. 
        if not not blankspace:                                      #If the list blankspace is not empty, then there is a column named ' ' that exists. 
            dfname.drop(' ', axis = 1, inplace = True)              #If the name of the column is ' ', drop it. These columns are blank.
        else:
            pass
        
        startcols = [col for col in dfname.columns if col == 'Start']   #This list will be empty unless there is a column named 'Start'
        if not not startcols:                                           #If the list is not empty, drop those column. 
            dfname.drop('Start', axis = 1, inplace = True)              #You need the portion creating a list to confirm that the column exists.           
        else:                                                           #If the column doesnt exist and you call drop, it will raise an error.
            pass
        
        typeindex = [idx for idx, val in enumerate(dfname.columns) if val == 'Type']  #Create a list of index values where the column name is 'Type'
        if len(typeindex) > 1:                  #For some machines, there are two parameters called 'Type'. We only want the second, so we will rename the other Type2
            newcols = dfname.columns.tolist()   #If there is only one parameter called 'Type', then it will skip this section. 
            num = typeindex[0]                  #Duplicate column names will force pd.concat to fail
            newcols[num] = 'Type2'              #Replace the name
            dfname.columns = newcols            #Rename the columns
    
        speedindex = [idx for idx, val in enumerate(dfname.columns) if val == 'Speed']  #Same thing as above with speed
        if len(speedindex) > 1:                 #If there is more than one named 'Speed', chage the second 'Speed' and rename the column.
            newcols = dfname.columns.tolist()
            num = speedindex[1]
            newcols[num] = 'Speed2'
            dfname.columns = newcols
        
    # =============================================================================
    # Read files (must be xlsx file)
    # =============================================================================
    
    with pd.ExcelWriter(finalpath) as writer:                   #Setup ExcelWriter to write the results to separate sheets in the same spreadsheet
        collectdf = pd.DataFrame(columns = originalheaders) #Create an empty dataframe to collect the reformatted data in. We will use this to concat later on
        df = pd.read_excel(exportfile, header = None)       #load the excel data into a dataframe
        
        df = pd.concat([pd.DataFrame(data = [[np.nan,np.nan]]), df, pd.DataFrame(data = [[np.nan,np.nan]]), pd.DataFrame(data = [[np.nan,np.nan]])], ignore_index = True)
    
    #Collection dataframe
    collectdf = pd.DataFrame(columns = originalheaders)
    #Find category labels (trauma, head, chest, etc)
    categories = []
    for row in df.index[:-6]:
        if (pd.isna(df.loc[row, 0])) & (pd.isna(df.loc[row+2, 0])) & (pd.isna(df.loc[row+4, 0])) & (pd.isna(df.loc[row+6, 0])):
            categories.append(row+1)
    #add last row so we have a stopping point for the last category
    categories.append(df.index[-1])
    
    #Isolate each category
    for c_idx, c in enumerate(categories[:-1]):
        catdf = df.loc[c:categories[c_idx +1]-2, :].copy()
        catdf = pd.concat([catdf, pd.DataFrame(data = [[np.nan,np.nan]]), pd.DataFrame(data = [[np.nan,np.nan]])], ignore_index = True)
    
        print(catdf)
        #store the category name
        category = catdf.iloc[0,0]
        #get rid of the first two rows (contains the category name and a blank line)
        #now the first line is the first protocol name
        catdf = catdf.iloc[2:, :]
    
        #find each protocol name
        protocols = []
        for row in catdf.index[:-4]:
            if (pd.isna(catdf.loc[row+1, 0])) & (pd.isna(catdf.loc[row+3, 0])) & (pd.isna(catdf.loc[row, 1])):
                protocols.append(row)
        protocols.append(catdf.index[-1])
        
        #Isolate each protocol
        for p_idx, p in enumerate(protocols[:-1]):
            protdf = catdf.loc[p:protocols[p_idx+1]-2, :].copy()
            protdf = pd.concat([protdf, pd.DataFrame(data = [[np.nan,np.nan]]), pd.DataFrame(data = [[np.nan,np.nan]])], ignore_index = True)
    
            print(protdf)
            protocol = protdf.iloc[0,0]
            protdf = protdf.iloc[2:,:]
            
            #find each phase
            acquisitions = []
            for phaserow in protdf.index:
                if 'acquisition label' in str(protdf.loc[phaserow, 0]).lower():
                    acquisitions.append(phaserow)
            acquisitions.append(protdf.index[-1])
            
            
            for a_idx, a in enumerate(acquisitions[:-1]):
                acqdf = protdf.loc[a:acquisitions[a_idx+1]-2, :].copy()
                # acqdf = pd.concat([acqdf, pd.DataFrame(data = [[np.nan,np.nan]]), pd.DataFrame(data = [[np.nan,np.nan]])], ignore_index = True)
                recons = []
                for acqrow in acqdf.index:
                    if 'result label' in str(acqdf.loc[acqrow, 0]).lower():
                        recons.append(acqrow)
                
                #If there are no recons, move on
                if not recons:
                    print(acqdf)
                    #temporary df to collect this scan's details
                    tempdf = pd.DataFrame(columns = originalheaders)
                    tempdf.loc[0, 'Protocol'] = protocol 
                    
                    acqname = acqdf.iloc[0,0]
                    acqname = acqname.split(':')[1].strip()
                    tempdf.loc[0, 'Scan/Recon Type'] = acqname 
                    #name of acquisition label
                    
                    for acqidx in acqdf.index[2:]:
                        param = acqdf.loc[acqidx, 0]
                        value = acqdf.loc[acqidx, 1]
                        if param in tempdf.columns:
                            tempdf.loc[0, param] = value
                    collectdf = pd.concat([collectdf, tempdf], ignore_index = True)
                            
                else:
                    recons.insert(0, acqdf.index[0])
                    recons.append(acqdf.index[-1]+2)
                    if recons[-1] == catdf.index[-1]:
                        recons[-1] = recons[-1]+2
                    #separate scans from recons
                    for r_idx, r in enumerate(recons[:-1]):
                        recondf = acqdf.loc[r:recons[r_idx+1]-2, :].copy()
                        # print(recondf)
                        #First is the scan parameters. Everything after are recons
                        tempdf = pd.DataFrame(columns = originalheaders)
                        tempdf.loc[0, 'Protocol'] = protocol 
                        if r_idx == 0:
                            acqname = recondf.iloc[0,0]
                            acqname = acqname.split(':')[1].strip()
                        else: 
                            acqname = recondf.iloc[0,0]
                            acqname = acqname.split(':')[1].strip()
                            acqname = 'Recon ' + acqname
                        tempdf.loc[0, 'Scan/Recon Type'] = acqname 
                        
                        for acqidx in recondf.index[2:]:
                            param = recondf.loc[acqidx, 0]
                            value = recondf.loc[acqidx, 1]
                            if param in tempdf.columns:
                                tempdf.loc[0, param] = value
                        collectdf = pd.concat([collectdf, tempdf], ignore_index = True)
                            
    
    
                            
            
    for f1, f2 in zip(originalheaders, finalheaders):       #Change the GE-specific names (without the leading and trailing whitespace) to our chosen names. 
        collectdf = collectdf.rename(columns = {f1:f2})
    
    #Add in scout angle
    for row in collectdf.index:
        if pd.isna(collectdf.loc[row, 'View Angle']):
            pass
        else:
            collectdf.loc[row, 'Scan/Recon Type'] = str(collectdf.loc[row, 'Scan/Recon Type'])+ ' (' +str(collectdf.loc[row, 'View Angle']) +')'
    
        if pd.isna(collectdf.loc[row, 'IR']):
            pass
        else:
            collectdf.loc[row, 'IR'] = 'iDose = '+ str(collectdf.loc[row, 'IR'])
            
    
        
        #Calculate mA from effective mAs
        if pd.isna(collectdf.loc[row, 'mAs']):
            pass
        else:
            if pd.isna(collectdf.loc[row, 'Pitch']):
                collectdf.loc[row, 'mA'] = int(float(collectdf.loc[row, 'mAs'])/float(collectdf.loc[row, 'Rot (s)']))
            else:
                collectdf.loc[row, 'mA'] = int(float(collectdf.loc[row, 'mAs'])*float(collectdf.loc[row, 'Pitch'])/float(collectdf.loc[row, 'Rot (s)']))
    
    collectdf.drop('View Angle', axis = 1, inplace = True)
            
    collectdf.drop('mAs', axis = 1, inplace = True)
    
    for row in collectdf.index[1:]:
      if 'Recon' in collectdf.loc[row, 'Scan/Recon Type']:                    #If the row is for recon settings, 
            try:
                if 'Recon' not in collectdf.loc[row-1, 'Scan/Recon Type']:      #If the previous row is not a Recon row (so it is a scan row)
                    reconset = ['Thick', 'Int', 'DFOV', 'Kernel', 'IR'] #The Int value will also be replaced (some GE machines have an Int val of 80 instead of 2.5)
                    
                    collectdf.loc[row-1, reconset] = collectdf.loc[row, reconset]   #Add in the values from the recon row to the row above
                    collectdf.loc[row-1, 'Scan/Recon Type'] = str(collectdf.loc[row-1, 'Scan/Recon Type']) + r'/' + str(collectdf.loc[row, 'Scan/Recon Type']) #Combine the scan/recon types together to reflect the new data in the row.
                    collectdf.drop(row, axis = 0, inplace = True)               #Drop the recon row.
                else:                                                           #If the previous row is a recon row, then the row can stay the same on its own. 
                  pass
            except KeyError:
                pass
      else:                                                                   #If this isn't a recon row, leave it alone
          pass
         
            
    collectdf.to_excel(finalpath, sheet_name = machinename, index = False)           
            
    finaldesname = 'Print_Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    
    despath = os.path.join(dirname, finaldesname)
        
    PDFformatter.PDF_formatter(finalpath, despath)
    
    
            
            
            
        
        