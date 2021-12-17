# =============================================================================
# Files must be filed: 'Main Folder'\'Machine name containing the vendor name (GE, Toshiba, or Siemens)\2020\file
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
from GE_settings import finalheaders, keyword1, keyword2, originalheaders
import PDFformatter



# =============================================================================
#Define Functions
# =============================================================================

# =============================================================================
# clean_function removes blank and duplicate columns from the df to allow concatenation
# =============================================================================
def GE(filename, machinename, typeofscanner):
    foldpath = r'Z:\Emi\Prots'

    dirname = os.path.dirname(filename)

    finalfilename = 'Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    
    finalpath = os.path.join(dirname, finalfilename)
    
    year = dt.datetime.today().strftime('%Y')
    
    GEpaths = [filename]
    
    

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
        for sheetidx, exportfile in enumerate(GEpaths):         #Iterate through the GE paths
            collectdf = pd.DataFrame(columns = originalheaders) #Create an empty dataframe to collect the reformatted data in. We will use this to concat later on
            df = pd.read_excel(exportfile, header = None)       #load the excel data into a dataframe
            
            df = df.loc[df[0] != 'AutoTransfer:']
            
            df1 = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)    #This section creates a data frame with one blank row and adds it to df.
            df = df.append(df1, ignore_index=True)                              #The GE format needs 3 extra blank rows at the end of the df in order to properly extract the last protocol.
            df = df.append(df1, ignore_index=True)
            df = df.append(df1, ignore_index=True)
            
            examidx = df.loc[df[0] == keyword1].index.tolist()          #Find what rows indicate the start of a protocol. keyword1 == 'Exam Dose Settings'. Get a list of the indices.
            
            examnamelist = []                                           #Empty list to collect the exam names
            for row in examidx:                                         
                examname = ''                                           #Empty string to build the name of each exam
                for q in range(0,4):                                    #Exam name text is contained in the first 4 columns of the GE data, depending on the machine
                    if pd.notna(df.iloc[row-2, q]):                     #Add the text to the string if the cell is not blank. The name of the protocol is two rows up from the keyword. 
                        examname = examname + str(df.iloc[row-2, q])    
                    else:
                        pass
                examname = ' '.join(examname.split())
                examnamelist.append(examname)
            
            examidx.append(df.shape[0]-1)                               #We need a stopping point for the last protocol in the data, so we append the last index point. 
                                                                        #Do this after pulling the exam names so that we don't pull a blank exam name. 
            for j in range(0, len(examidx)-1):
                startidx = examidx[j]                                   #Starting point for the protocol
                stopidx = examidx[j+1]                                  #Stopping point for the protocol
                protocoldf = df.loc[startidx:stopidx, :]                #Pull the data only for one protocol. 
                seriesloc = protocoldf.loc[protocoldf[0].str.contains(keyword2, na = False)]    #Find rows that contain the keyword, indidcate the start of each scan or recon settings. keyword2 = 'Series'
                seriesidx = seriesloc.index.tolist()                                            #Write the index values to a list.
                seriesidx.append(protocoldf.index[-2]-1)                                        #Add a stopping point for the last group of settings in the protocol
                # =============================================================================
                #  Create single rows for each scan or recon in the series.       
                # =============================================================================
                for i in range(0, len(seriesidx)-1):
                    seriesname = df.loc[seriesidx[i],0]                 #Pull the name of the series (e.g. Series 2 Group 1 Scan Settings)
                    startidx_ser = seriesidx[i] + 1                     #Starting point for series
                    stopidx_ser = seriesidx[i+1] - 1                    #Stopping point for series
                    seriesxdf = df.loc[startidx_ser:stopidx_ser, :]     #Isolate the data for just that series
                    seriesxdf = seriesxdf.loc[seriesxdf[0].notna()]     #Drop blank rows
                    if seriesxdf.shape[0] > 2:                          #There will either be 2, 4, or 5 rows of data
                        seriesxdf3 = seriesxdf.iloc[2:, :]              #The series that have more than 2 rows have irrelevant parameters in the first 2 rows. Drop them.
                        seriesxdf3 = seriesxdf3.dropna(axis = 1)        #Drop any columns that are completely blank.
                        seriesxdf3 = seriesxdf3.reset_index()           #Reset the index
                        seriesxdf3 = seriesxdf3.drop('index', axis = 1) #pandas automatically adds an additional index column when you reset the index, drop it.
            
                        new_header = seriesxdf3.iloc[0].tolist()        #The first row contains what we want our column headers to be.
                        for idx, name in enumerate(new_header):         #Iterate through the names in the first row
                            if name != ' ':                             #If the name is NOT a blank space (this exists in GE data). We will drop it later.
                                new_header[idx] = str(name).strip()     #Strip the name of any whitespace and replace the name in the list. 
                            else:
                                pass
            
                        seriesxdf3.columns = new_header                 #Replace the column names in the dataframe
                        seriesxdf3 = seriesxdf3.drop(0, axis = 0)       #Now that we've renamed the columns, we can drop the first row of data that contains the column names  
                    
                    else:                                               #If its only 2 rows (1st row = column names, 2nd row = data)
                        seriesxdf3 = seriesxdf.dropna(axis = 1)         #Drop any blank columns
                        seriesxdf3 = seriesxdf3.reset_index()           #Reset the index
                        seriesxdf3 = seriesxdf3.drop('index', axis = 1) #Drop the extra index column it creates
                        
                        new_header = seriesxdf3.iloc[0].tolist()        #Write the first row (column names) to a list
                        for idx, name in enumerate(new_header):         #Iterate through
                            if name != ' ':                             #Remove the whitespace if the name is not ' '
                                new_header[idx] = str(name).strip()
                            else:
                                pass
                        
                        seriesxdf3.columns = new_header                 #Replace the column names in the dataframe
                        seriesxdf3 = seriesxdf3.drop(0, axis = 0)       #Drop the first row
                    
                    if seriesxdf3.shape[0] > 2:                         #Now, there should be either 1 or 2 lines of data in the df. The only reason there is 2 is if there are 2 scouts in 
                        for m in range(0, seriesxdf3.shape[0]-1):       #in Series 1. In that case, we need to separate it into two separate dataframes.
                            seriesxdf3_m = pd.DataFrame(data = [seriesxdf3.iloc[m]], columns = seriesxdf3.columns)  #Write each row into its own dataframe
            
                            clean_function(seriesxdf3_m)                                            #This function cleans up the data and removes the columns that cause issues with concatenation. 
            
                            seriesxdf3_m.insert(0, 'Scan/Recon Type', seriesname)                   #Insert the Scan/Recon Type as a new column       
                            seriesxdf3_m.insert(0, 'Protocol', examnamelist[j])                     #Add the protocol name as a new column
                            collectdf = pd.concat([collectdf, seriesxdf3_m], ignore_index= True)    #The row is not complete. Add this row to the collection datafrane.
            
                    else:                                               #If there is only one row of data, clean the data, add the new columns, and add the row to the collectiondf.
                        clean_function(seriesxdf3)
                        
                        seriesxdf3.insert(0, 'Scan/Recon Type', seriesname)             
                        seriesxdf3.insert(0, 'Protocol', examnamelist[j])
                        collectdf = pd.concat([collectdf, seriesxdf3], ignore_index= True)
            # =============================================================================
            #  Calculate pitch   
            # =============================================================================
            halfpitch = []                                                  #Empty list to collect rows that need pitch cut in half so we can adjust the collimation value 
            if not pd.isna(collectdf['pitch']).all():                       #If the pitch column is not completely blank
                for row in collectdf.index:                                     
                    if pd.notna(collectdf.loc[row, 'pitch']):               #If there is a value in the pitch cell, 
                        rowsval = float(collectdf.loc[row, 'Rows'])         #Retrieve the value for Rows as a float
                        pitchfactor = float(collectdf.loc[row, 'pitch'])    #Get the pitch value as a float
                        pitchval = pitchfactor/rowsval                      #Divide pitch by # of rows to get pitch factor
                        if pitchval > 1.5:                                  #Divide any pitch values greater than 1.5 by 2
                            pitchval = pitchval/2
                            halfpitch.append(row)                           #Append the row to the halfpitch list
                        else:
                            pass
    
                        collectdf.loc[row, 'pitch'] = pitchval              #Replace with the new pitch value
                    else:
                        pass
            if 'Speed2' in collectdf.columns:                               #If there is a speed2 value (this is the table speed). There is a 'pitch' column for machines without a 'Speed2' column
                for row in collectdf.index:                             
                    if pd.notna(collectdf.loc[row, 'Speed2']):              #If there is a value for table speed   
                        speedval = float(collectdf.loc[row, 'Speed2'])      #Define table speed, convert to float
                        rowsval = float(collectdf.loc[row, 'Rows'])         #Define number of rows, convert to float
                        pitchval = speedval/(rowsval*0.625)                 #Calculate pitch value
                        if pitchval > 1.5:                                  #Divide any large pitch values by 2
                            pitchval = pitchval/2
                            halfpitch.append(row)                           #Record the row so coll can be updated
                        else:
                            pass
                        collectdf.loc[row, 'pitch'] = pitchval              #Add pitch value to dataframe
                    else:
                        pass
            else:
                pass
            # =============================================================================
            #  Replace minma/maxma for ECG  
            # =============================================================================
            for row in collectdf.index:
                if pd.notna(collectdf.loc[row, 'ECGMinmA']):
                    collectdf.loc[row, 'MinmA'] = collectdf.loc[row, 'ECGMinmA']
                else:
                    pass
                if pd.notna(collectdf.loc[row, 'ECGMaxmA']):
                    collectdf.loc[row, 'MaxmA'] = collectdf.loc[row, 'ECGMaxmA']
                else:
                    pass
                
            # =============================================================================
            #  Rename the columns with our generic names   
            # =============================================================================
            collectdf = collectdf[originalheaders]                  #Trim the data to only columns we are interested in
            
            for f1, f2 in zip(originalheaders, finalheaders):       #Change the GE-specific names (without the leading and trailing whitespace) to our chosen names. 
                collectdf = collectdf.rename(columns = {f1:f2})
            # =============================================================================
            #  Reformat the protocol names
            # =============================================================================
            collectdf['Protocol'].replace({' +':' '},inplace = True, regex=True)
            # =============================================================================
            #  Drop any blank rows and classify scans and recons
            # =============================================================================
            collectdf.drop(['ECGMinmA', 'ECGMaxmA'], axis = 1, inplace = True)
    
            for row in collectdf.index:                           
                blankfields = finalheaders[2:]                      #We only want to look at the rows after Protocol and Scan/Recon Type, since there is always a value in those columns
        
                if pd.isna(collectdf.loc[row, blankfields]).all():  #If the row past the first two columns is completely blank, drop it.
                    collectdf.drop(row, axis = 0, inplace = True)   #Drop the row.
                    continue                                        #Break out of the loop for this row
                if pd.notna(collectdf.loc[row, 'Plane']):               #If the row has a value in the 'Plane' column, it is a scout scan.
                    collectdf.loc[row, 'Scan/Recon Type'] = 'Scout'     #Replace the name with Scout
                    continue                                            
                if 'Recon' in collectdf.loc[row, 'Scan/Recon Type']:    #If the row has the word 'Reocn' in the 'Scan/Recon Type' column, it is a Recon
                    collectdf.loc[row, 'Scan/Recon Type'] = 'Recon'     #Replace the name with Recon
                    continue
                if 'Scan' in collectdf.loc[row, 'Scan/Recon Type']:     #If the row has the word 'Scan in the the 'Scan/Recon Type' column, it is a scan
                    collectdf.loc[row, 'Scan/Recon Type'] = 'Scan'      #Replace the name with Scan
                    continue
                else:                                                   #Otherwise, pass. The remaining will be in the format 'Series #'
                    pass
            # # =============================================================================
            # # Identify SmartPrep scans from the 'SmartPrep' column
            # # =============================================================================
            # for row in collectdf.index:                                 #Go through the rows again now that they are renamed:
            #     try:                                                    #Try because code will encounter an error if celles in the 'SmartPrep' column are blank.
            #         if ('false' in collectdf.loc[row, 'SmartPrep']) | ('No' in collectdf.loc[row, 'SmartPrep']):    #If it says Smartprep is off, drop the row. There is no other useful info in these rows. 
            #             collectdf.drop(row, axis = 0, inplace = True)                                               
            #             continue                                                                                    #Break out of the loop for this row. 
            #         if ('true' in collectdf.loc[row, 'SmartPrep']) | ('Yes' in collectdf.loc[row, 'SmartPrep']):    #If SmartPRep is on, rename the next row from scan to SmartPrep
            #             collectdf.loc[row+1, 'Scan/Recon Type'] = 'SmartPrep'
            #             collectdf.drop(row, axis = 0, inplace = True)                                               #Drop the row. 
            #             continue                                                                                    #Break out of the loop for this row.
            #         else:                                                                                           #Otherwise, pass.
            #             pass
            #     except TypeError:       #Pass if it encounters an error.
            #         pass
            # collectdf.drop('SmartPrep', axis = 1, inplace = True)       #Drop the column 'SmartPrep' now that we have what we need from it.
            # =============================================================================
            # Add in Axial or Helical into the scan/recon type from the 'Type' column
            # =============================================================================
            for row in collectdf.index:                 
                if pd.isna(collectdf.loc[row, 'Type']):         #If the type column is blank, pass
                    pass
                else:                                           #Otherwise, add the Test from the type column to the Scan/Reocn Type column.
                    collectdf.loc[row,'Scan/Recon Type'] = collectdf.loc[row,'Type'] + ' ' + collectdf.loc[row, 'Scan/Recon Type']
            collectdf.drop('Type', axis = 1, inplace = True)    #Drop the column 'Type now that we have what we need from it.
    
            # =============================================================================
            # Combine the recon parameters into the corresponding scan row. 
            # =============================================================================
            for row in collectdf.index:
                if 'Recon' in collectdf.loc[row, 'Scan/Recon Type']:                    #If the row is for recon settings, 
                    try:                                                                #Use try so avoid key error
                        if 'Recon' not in collectdf.loc[row-1, 'Scan/Recon Type']:      #If the previous row is not a Recon row (so it is a scan row)
                            if pd.isna(collectdf.loc[row-1, 'Thick']):                  #If it does not already have a value for Thickness, use the recon thickness
                                reconset = ['Thick', 'Int', 'DFOV', 'Kernel', 'IR', 'IQEnhance'] #The Int value will also be replaced (some GE machines have an Int val of 80 instead of 2.5)
                            else:                                                       #If the scan row already has a value for thickness, keep that value
                                reconset = ['DFOV', 'Kernel', 'IR', 'IQEnhance']        #Same parameters as above without the thickness
                            collectdf.loc[row-1, reconset] = collectdf.loc[row, reconset]   #Add in the values from the recon row to the row above
                            collectdf.loc[row-1, 'Scan/Recon Type'] = str(collectdf.loc[row-1, 'Scan/Recon Type']) + r'/' + str(collectdf.loc[row, 'Scan/Recon Type']) #Combine the scan/recon types together to reflect the new data in the row.
                            collectdf.drop(row, axis = 0, inplace = True)               #Drop the recon row.
                        else:                                                           #If the previous row is a recon row, then the row can stay the same on its own. 
                            pass
                    except KeyError:                                                    #If there is an error, pass
                        pass
                else:                                                                   #If this isn't a recon row, leave it alone
                    pass
            # =============================================================================
            # Add in the detector thickness into the Collimation column and axial
            # =============================================================================
            for row in collectdf.index:
                if (pd.notna(collectdf.loc[row, 'Coll'])) & (collectdf.loc[row, 'Coll'] != ' SmartCollimation'):    #If the  Collimation column is not blank and does not equal SmartCollimation
                    if row in halfpitch:                                                                        #Change the coll value if we divided pitch by 2
                        collectdf.loc[row, 'Coll'] = str(collectdf.loc[row, 'Coll']) + ' x 1.250 mm' 
                    else:
                        collectdf.loc[row, 'Coll'] = str(collectdf.loc[row, 'Coll']) + ' x 0.625 mm'        #Add in x0.625 mm to the value in the cell
                else:                                                                                   #Otherwise, leave it alone
                    pass
                if (pd.notna(collectdf.loc[row, 'Mode'])) & ('Axial' in collectdf.loc[row, 'Scan/Recon Type']): #Int = thickness for axial scans
                    mode = float(collectdf.loc[row, 'Mode'][:-1])                                               #if mode*thickness != interval 
                    thickness = float(collectdf.loc[row, 'Thick'])
                    try:
                        interval = float(collectdf.loc[row, 'Int'])
                    except ValueError:
                        interval = float(str(collectdf.loc[row, 'Int']).split(' ')[0])
                    product = mode*thickness
                    if product != interval:
                        collectdf.loc[row, 'Int'] = collectdf.loc[row, 'Thick']
                    else:
                        pass
                else:
                    pass
            # =============================================================================
            # Add in the Plane value into the Scan/Recon Type
            # =============================================================================
            for row in collectdf.index:
                if pd.notna(collectdf.loc[row, 'Plane']):                                               #If there is a value in the 'Plane' columns, add the plane value into the scan/recon type (It will always be a scout scan)    
                    collectdf.loc[row, 'Scan/Recon Type'] = str(collectdf.loc[row, 'Scan/Recon Type']) + ' ' + str(collectdf.loc[row, 'Plane'])
                else:
                    pass
            collectdf.drop(['Plane', 'Mode'], axis = 1, inplace = True)                 #Drop the Plane column now that we have what we need.
                    
            collectdf.to_excel(writer, sheet_name = machinename, index=False)           #Write the data to excel, name the sheet as the machine name
            print(machinename)                                                          #Print the name of the machine
    
    finaldesname = 'Print_Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    despath = os.path.join(dirname, finaldesname)
    
    PDFformatter.PDF_formatter(finalpath, despath)

            
            
            
            
            