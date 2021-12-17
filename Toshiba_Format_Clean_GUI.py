# -*- coding: utf-8 -*-
# =============================================================================
# XML file must be in ***Protocol Downloads\DirectoryforMachine\2020\file.xml
# The XML file must be the ONLY file in the 2020 folder
# =============================================================================

"""
Created on Wed Sep 16 11:28:40 2020

@author: EastmanE
"""

import os
import datetime as dt
import pandas as pd
import openpyxl as xl
import numpy as np
import xml.etree.ElementTree as ET
from Toshiba_settings import fields, renamefields, finalfields2, MOTfields, MOTfinalfields2, reconfields
import PDFformatter

foldpath = r'Z:\Emi\Prots'
finalpath = r'Z:\Emi\Prots\Toshiba.xlsx'   #Path to write the formatted protocols
year = dt.datetime.today().strftime('%Y')


# =============================================================================
# Read all Toshiba files (must be xml file)
# =============================================================================

    
    

def Toshiba(filename, machinename, typeofscanner):
    foldpath = r'Z:\Emi\Prots'

    dirname = os.path.dirname(filename)

    finalfilename = 'Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    reconfilename = 'ToshibaRecon_' +dt.datetime.now().strftime("%Y%m%d%H%M%S")+ '.xlsx'
    finalpath = os.path.join(dirname, finalfilename)
    finalreconpath = os.path.join(dirname,reconfilename)
    
    year = dt.datetime.today().strftime('%Y')
    
    toshpaths = [filename]

    with pd.ExcelWriter(finalpath) as writer:                   #Setup ExcelWriter to write the results to separate sheets in the same spreadsheet
        with  pd.ExcelWriter(finalreconpath) as reconwriter:    
            for sheetidx, exportfile in enumerate(toshpaths):           #Iterate through the Toshiba/Canon paths
                tree = ET.parse(exportfile)                             #Use XML etree to get the data. Get tree 
                root = tree.getroot()                                   #Get roots
                
                taglist = [elem.tag for elem in root.iter()]            #Iterate through the roots and get the tags of each element (category). Write the keywords to a list
                textlist = [elem.text for elem in root.iter()]          #Iterate through the roots and pull the text associated with each tag (value). Write the values to a list
    
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                    suplist = [c.findall(".//sup")[0].tail for c in root.findall(".//mA")]
                    for idx, i in enumerate(textlist):
                        if taglist[idx] == 'mA':
                            textlist[idx] = suplist[0]
                            del suplist[0]
            
    
    
    
                df = pd.DataFrame(data = taglist)                       #Create a dataframe from the two lists. Parameters will be listed in the first columns.
                df[1] = textlist                                        #Values are written into the second column.
                
                # =============================================================================
                # Break up the data based on 'PatientType' 
                # =============================================================================
                if 'Genesis' in typeofscanner:
                    patienttype = []                                                                    #Blank list to collect the index values of the first section breaks.
                    for row in df.index:                                    
                        if (df.loc[row, 0] == 'PatientType') & (df.loc[row,1] == None):                 #Sections are broken up by 'PatientType' tag with no associated text. Other 'PatientType' tags have values with them. 
                            if (df.loc[row-1, 0] == 'PatientTypeName') & (df.loc[row-1, 1] != None):    #There is one instance where the PatientType value is misplaced as 'PatientTypeName'
                                df.drop(row, axis = 0, inplace = True)                                  #Drop the blank 'PatientType' row and use the index of the 'PatientTypeName' row.
                                patienttype.append(row-1)                                               #Add the index value to the list.
                            else:
                                patienttype.append(row)                                                 #If it is not the scenario above, then append the index value.
                                
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                    patienttype = []
                    for row in df.index:
                        if (df.loc[row, 0] == 'DefSureIQParamHeading') | (df.loc[row,0] == 'DefSureExpReportHeading')| (df.loc[row,0] == 'BodyTypeName'):
                            patienttype.append(row)
                patienttype.append(df.index[-1])
                # =============================================================================
                # Pull SureIQ Settings contained in df[patienttype[0]:patienttype[2]] 
                # =============================================================================
                final_sureIQdf = pd.DataFrame()                             #Create a blank df to collect sureIQ settings
                if 'Genesis' in typeofscanner:    
                    startnum = 0
                    stopnum = 2
                    columndrop = ['DefParams']
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                    startnum = 1
                    stopnum = 2
                    columndrop = ['DefParams', 'RowSpan', 'DefSureExpParam']
                for ii in range(startnum, stopnum):
                    IQstartidx = patienttype[ii]                            #Define starting index
                    IQstopidx = patienttype[ii+1]                           #Define stopping index
                    sureIQdf= df[IQstartidx:IQstopidx]                      #Isolate sureIQ data 
                    
                    for row in sureIQdf.index:                              #Sections are broken up bythe tag 'SureAnatomical'. There is always a 'DefParam' tag in the row below that is also blank.
                        if sureIQdf.loc[row, 0] == 'SureAnatomical':        #Drop the 'SureAnatomical' rows because they don't contain any useful info. 
                            sureIQdf.drop(row, axis = 0, inplace = True)    #We will use the tag 'DefParam' to identify where to break up the sections. 
                        else:
                            pass
                    
                    defparams= []                                           #Empty list to collect index values for section breaks based on 'DefParams'
                    for row in sureIQdf.index:                              #Find index values.
                        if sureIQdf.loc[row, 0] == 'DefParams':
                                defparams.append(row)
                        else:
                            pass
                    defparams.append(IQstopidx)                             #Append the index value of the last row so that there is a stopping point for the last section of data.
                    
                    for i in range(0, len(defparams)-1):                    #Iterate through the 'DefParams' index values
                        sureIQstart = defparams[i]                          #Define starting point for the section
                        sureIQstop = defparams[i+1] - 1                     #Define stopping point
                        single_sureIQdfcols = sureIQdf.loc[sureIQstart:sureIQstop, 0].tolist()                      #first column values with be the column headers. Write is to a list
                        single_sureIQdfvals = [sureIQdf.loc[sureIQstart:sureIQstop, 1].tolist()]                    #Second column values will the the data. make sure the list is in a list so that it writes to a row later in the df
                        single_sureIQdf = pd.DataFrame(data = single_sureIQdfvals, columns= single_sureIQdfcols)    #Create a dataframe with the sureiq settings- this is only one line.
                        try:
                            single_sureIQdf.drop('DefSureExpParams', axis = 1, inplace = True)
                        except KeyError:
                            pass
                        final_sureIQdf = pd.concat([final_sureIQdf, single_sureIQdf], ignore_index = True)          #Add this as a row to our collection df
                
                try:
                    final_sureIQdf.drop(columndrop, axis=1, inplace = True)
                    final_sureIQdf['SD lookup'] = final_sureIQdf['Organ']+ ' ' + final_sureIQdf['Name']             #Create a lookup value that we can use to pull data into our protocols based on their SUreIQ setting later on.
                    final_sureIQdf = final_sureIQdf.rename(columns = {'SliceThickness':'SD Thick'})                 #Rename one column to match our naming convention.
                except KeyError:
                    pass
                
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                        # final_sureIQdf['SD lookup'] = final_sureIQdf['DefSureExpName']+ ' ' + final_sureIQdf['DefSureExpSureIQ']             #Create a lookup value that we can use to pull data into our protocols based on their SUreIQ setting later on.
                        final_sureIQdf = final_sureIQdf.rename(columns = {'DefSureExpSliceThickness':'SD Thick'})                 #Rename one column to match our naming convention.
                        final_sureIQdf = final_sureIQdf.rename(columns = {'DefSureExpSD':'SD'})                 #Rename one column to match our naming convention.
                        final_sureIQdf = final_sureIQdf.rename(columns = {'DefSureExpMaxmA':'MaxmA'})                 #Rename one column to match our naming convention.
                        final_sureIQdf = final_sureIQdf.rename(columns = {'DefSureExpMinmA':'MinmA'})                 #Rename one column to match our naming convention.
        
                # final_sureIQdf.to_excel(writer, sheet_name = 'SureIQ- ' + machinename, index=False)           #Write the data to excel, name the sheet as the machine name
                
                # =============================================================================
                # Pull RECON Settings    
                # =============================================================================
                final_RECONdf = pd.DataFrame()                          #Create an empty df to collect the recon settings.
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                    RECONstartidx = patienttype[0]                          #Define starting point
                    RECONstopidx = patienttype[1]                           #define stopping point
                else:
                    RECONstartidx = patienttype[2]                          #Define starting point
                    RECONstopidx = patienttype[3]                           #define stopping point
                RECONdf= df[RECONstartidx:RECONstopidx]                 #Isolate the data
                
                for row in RECONdf.index:
                    if RECONdf.loc[row, 0] == 'Anatomical':             #Sections are broken up by the tag'Anatomical'. 'DefParams' is directly under each term. 
                        RECONdf.drop(row, axis = 0, inplace = True)     #Drop the 'Anatomical' rows so that we can index by 'DefParams'
                    else:
                        pass
                
                RECONdefparams= []                              #Blank list to collect index values
                for row in RECONdf.index:                       
                    if RECONdf.loc[row, 0] == 'DefParams':      #Write index values of rows containing 'DefParams' to the list
                            RECONdefparams.append(row)
                    else:
                        pass
                RECONdefparams.append(RECONstopidx)             #Add a stopping point for the final set of data
            
                for i in range(0, len(RECONdefparams)-1):       #Iterate through the index values. These new sections will each become one row. 
                    RECONstart = RECONdefparams[i]              #Define a starting point
                    RECONstop = RECONdefparams[i+1] - 1         #Define a stopping point
                
                    single_RECONdfcols = RECONdf.loc[RECONstart:RECONstop, 0].tolist()                      #List of column headers from the first column in the df
                    single_RECONdfvals = [RECONdf.loc[RECONstart:RECONstop, 1].tolist()]                    #List of values from the second row in the df
                    single_RECONdf = pd.DataFrame(data = single_RECONdfvals, columns= single_RECONdfcols)   #Create the new df for the row
                    final_RECONdf = pd.concat([final_RECONdf, single_RECONdf], ignore_index = True)         #Add the new row to the collecting df.
                
                final_RECONdf.drop('DefParams', axis = 1, inplace = True)
        
                for colname in final_RECONdf.columns: 
                    if 'SureIQ' in colname:
                        newname = colname[6:]
                        final_RECONdf.rename(columns = {colname:newname}, inplace = True)
                # =============================================================================
                # Combine columns o create names
                # =============================================================================
                final_RECONdf.insert(0, 'SureIQ Name', '')  
                for row in final_RECONdf.index:
                    if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                        if pd.notna(final_RECONdf.loc[row, 'AnatomyName']):
                            final_RECONdf.loc[row, 'SureIQ Name'] = str(final_RECONdf.loc[row, 'AnatomyName']) + ' ' + str(final_RECONdf.loc[row, 'Selection'])
                        else:
                            final_RECONdf.loc[row, 'SureIQ Name'] = str(final_RECONdf.loc[row-1, 'AnatomyName']) + ' ' + str(final_RECONdf.loc[row, 'Selection'])
                            final_RECONdf.loc[row, 'AnatomyName'] = final_RECONdf.loc[row-1, 'AnatomyName']
                    else:
                        final_RECONdf.loc[row, 'SureIQ Name'] = str(final_RECONdf.loc[row, 'Organ']) + ' ' + str(final_RECONdf.loc[row, 'Name'])
                try:
                    final_RECONdf.rename(columns = {'ReconProcess':'AIDR'}, inplace = True)
                except KeyError:
                    pass
                final_RECONdf = final_RECONdf[reconfields]
                final_RECONdf.to_excel(reconwriter, sheet_name = machinename, index=False)           #Write the data to excel, name the sheet as the machine name
         
            # =============================================================================
            # Contrast Settings are contained in df[patienttype[3]:patienttype[4]] 
            # =============================================================================
        
                # =============================================================================
                # Pull Protocol Settings (contained in df[patienttype[4]:patienttype[5]])
                # =============================================================================
                final_PROTdf = pd.DataFrame() 
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:                         #MOT adult protocols are in 2:3, child and trauma protocols are in 4:6
                    PROTstartidx = patienttype[2]
                    PROTstopidx = patienttype[3]
                    PROTdf= df[PROTstartidx:PROTstopidx]
                    
                    childPROTstartidx = patienttype[4]
                    childPROTstopidx = patienttype[6]
                    PROTdf= pd.concat([PROTdf, df[childPROTstartidx:childPROTstopidx]], axis = 0)
        
                    for row in PROTdf.index:                    #Drop any rows that we don't need (not in the MOTfields list)
                        if PROTdf.loc[row, 0] in MOTfields:
                            pass
                        else:
                            PROTdf.drop(row, axis = 0, inplace = True)
                else:
                    PROTstartidx = patienttype[4]               #Protocols are in 4:7 for all other machines
                    PROTstopidx = patienttype[7]
                    PROTdf= df[PROTstartidx:PROTstopidx]
                    for row in PROTdf.index:                    #Drop rows we don't need
                        if PROTdf.loc[row, 0] in fields:
                            pass
                        else:
                            PROTdf.drop(row, axis = 0, inplace = True)
             
                examplannames = []                              #empty list to collect exam names
                for idx, row in enumerate(PROTdf.index):        #Add protocol names for the list
                    if PROTdf.loc[row, 0] == 'ExamPlanName':
                        examplannames.append(row)
                        
                        # =============================================================================
                        # Add Exam Number to Exam Name so they're all unique    
                        # =============================================================================
                        if PROTdf.loc[PROTdf.index[idx+1], 0] == 'ExamPlanNo':
                            examnum = str(PROTdf.loc[PROTdf.index[idx+1], 1]) 
                        elif PROTdf.loc[PROTdf.index[idx-1], 0] == 'ExamPlanNo':
                            examnum = str(PROTdf.loc[PROTdf.index[idx-1], 1]) 
                        else:
                            continue
                        if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                            examnum = examnum.split(': ', 1)[1]
                        else:
                            pass
                        
                        PROTdf.loc[row, 1] = str(PROTdf.loc[row, 1]).strip() + ' (' + examnum + ')'
                    else:
                        pass
                examplannames.append(PROTdf.index[-1])              #After you fix all of the names, append the index of the last row so there is a stopping point for the last protocol.  
        
                for jj in range(0, len(examplannames)-1):           #Break up into sections for each protocol
                    examstart = examplannames[jj]           
                    examstop = examplannames[jj+1] - 1
                    
                    examonly = PROTdf.loc[examstart:examstop]       #Only look at one protocol
                    
                    scanmodename= []                                #Find locations of each recon or scan
                    for row in examonly.index:                  
                        if (examonly.loc[row, 0] == 'ScanModeParam') | (examonly.loc[row, 0] == 'Reconstruction') | (examonly.loc[row, 0] == 'ScanoModeName'):
                                scanmodename.append(row)
                        else:
                            pass
                    scanmodename.append(examstop)                   #Append stopping point
            
                    for i in range(0, len(scanmodename)-1):         #Only look at one scan or recon
                        PROTstart = scanmodename[i]
                        PROTstop = scanmodename[i+1] - 1
                    
                        single_PROTdfcols = PROTdf.loc[PROTstart:PROTstop, 0].tolist()      #Turn first column into column headers
                        single_PROTdfcols.insert(0, 'Protocol Name')                        #Add in the Protocol Name column
                        single_PROTdfvals2 = PROTdf.loc[PROTstart:PROTstop, 1].tolist()     #Turn second column into df values
                        single_PROTdfvals2.insert(0, PROTdf.loc[examplannames[jj], 1])      #Add in the protocol name
                        single_PROTdfvals = [single_PROTdfvals2]                            
                        
                        single_PROTdf = pd.DataFrame(data = single_PROTdfvals, columns= single_PROTdfcols)  #Turn the lists into a df
                        final_PROTdf = pd.concat([final_PROTdf, single_PROTdf], ignore_index = True)        #Concatenate into the collection df
                        
                        if i > 0:                           #i>0 to avoid an index error
                            x = final_PROTdf.iloc[[-2]]     #Compare to second to last row in df
                            x = x.reset_index(drop = True)  #Reset index so we can concat if necessary
                        
                            try:                                                                                                #If there is a scan before the recon, add in 
                                if (pd.isna(x.loc[0,'Reconstruction'])) & (pd.notna(single_PROTdf.loc[0,'Reconstruction'])):    #the recon information to the scan row
                                    for colname in single_PROTdf.columns:                                                       
                                        if pd.notna(single_PROTdf.loc[0, colname]):
                                            x.loc[0,colname] = single_PROTdf.loc[0, colname]
                                        else:
                                            pass
                                    final_PROTdf.drop(final_PROTdf.tail(2).index,inplace=True)              #Drop the bottom 2 rows
                                    final_PROTdf = pd.concat([final_PROTdf, x], ignore_index= True)         
                    
                                else:
                                    pass
                            except KeyError:
                                pass
        
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:         #MOT has different keywords, replace them with the generic names
                    for typerow in final_PROTdf.index:
                        if pd.notna(final_PROTdf.loc[typerow, 'ScanoModeName']):
                            final_PROTdf.loc[typerow, 'ScanModeName'] = final_PROTdf.loc[typerow, 'ScanoModeName']
                        if pd.notna(final_PROTdf.loc[typerow, 'ScanokV']):
                            final_PROTdf.loc[typerow, 'kV'] = final_PROTdf.loc[typerow, 'ScanokV']
                        if pd.notna(final_PROTdf.loc[typerow, 'ScanomA']):
                            final_PROTdf.loc[typerow, 'mA'] = final_PROTdf.loc[typerow, 'ScanomA']
                        if pd.notna(final_PROTdf.loc[typerow, 'SXFC']):
                            final_PROTdf.loc[typerow, 'FC'] = final_PROTdf.loc[typerow, 'SXFC']
             
                    final_PROTdf.drop(['ScanModeParam', 'ScanoModeName', 'ScanokV', 'ScanomA', 'SXFC'], axis = 1, inplace = True)    #Drop these vendor specific column names
    
                else:
                    final_PROTdf['SD lookup'] = final_PROTdf['SureExpOrgan']+ ' ' + final_PROTdf['SureExpName']     #Use the lookup value to find the SD and current modulation settings
                    sureIQcolumns = final_sureIQdf[['SD lookup', 'SD', 'SD Thick', 'MinmA', 'MaxmA']]
                    final_PROTdf = final_PROTdf.merge(sureIQcolumns, on = 'SD lookup', how = 'left', suffixes = ['Recon', ''])    #merge them together side by side
                    final_PROTdf.drop('SD lookup', axis = 1, inplace = True)                                                        #Drop the lookup column
                   
                for row in final_PROTdf.index:
                    if pd.notna(final_PROTdf.loc[row, 'FC']):                                       #Add in 'FC' text to the FC value
                        final_PROTdf.loc[row, 'FC'] = 'FC ' + str(final_PROTdf.loc[row, 'FC'])
                    else:
                        pass
                    if pd.isna(final_PROTdf.loc[row, 'Reconstruction']):            
                        pass
                    else:                                                       #If the reconstruction is notempty and the scanmodename is empty, 
                        if pd.isna(final_PROTdf.loc[row, 'ScanModeName']):      #Replace the scanmodename with the Reconstruction value
                            final_PROTdf.loc[row,'ScanModeName'] = final_PROTdf.loc[row, 'Reconstruction']
                        else:                                                   #If Scanmodename is not empty and reconstruction is also not empty, combine the two with a slash
                            final_PROTdf.loc[row,'ScanModeName'] = str(final_PROTdf.loc[row,'ScanModeName']) + '/' + str(final_PROTdf.loc[row, 'Reconstruction'])
                final_PROTdf.drop('Reconstruction', axis = 1, inplace = True)   #Drop the reconstruction column 
            
            
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:                                         #For MOT, Add in a CTDI NV and SEMAR field. They are blank, but needed for consistency
                    final_PROTdf.insert(0, 'CTDI NV', '')
                    final_PROTdf.insert(0, 'SEMAR', '')
        
                    final_PROTdf= final_PROTdf[MOTfinalfields2]                 #Only keep columns that we need
                    for f1, f2 in zip(MOTfinalfields2, renamefields):           #Rename the fields with the generic rooms
                        final_PROTdf = final_PROTdf.rename(columns = {f1:f2})
                
                else:
                    final_PROTdf.insert(0, 'CTDI NV', '')                       
        
                    final_PROTdf= final_PROTdf[finalfields2]
                    for f1, f2 in zip(finalfields2, renamefields):
                        final_PROTdf = final_PROTdf.rename(columns = {f1:f2})
        
                final_PROTdf.insert(5, 'AEC IQ Ref', '')                        #Add in AEC IQ Ref Column and populate with the SD and SD thick vals
                
                if 'Toshiba/Canon Aquilion 64' in typeofscanner:
                    final_PROTdf.drop('SEMAR', axis = 1, inplace = True)        #Drop SEMAR, move min/max mA values
                    for row in final_PROTdf.index:
                        if (pd.notna(final_PROTdf.loc[row, 'SD'])) & (pd.notna(final_PROTdf.loc[row, 'SD Thick'])):
                            final_PROTdf.loc[row+1, 'AEC IQ Ref'] = str(final_PROTdf.loc[row, 'SD']) + ' @ ' + str(final_PROTdf.loc[row, 'SD Thick'])
                            final_PROTdf.loc[row+1, 'MinmA'] = final_PROTdf.loc[row, 'MinmA'] 
                            final_PROTdf.loc[row, 'MinmA'] = ''
                            final_PROTdf.loc[row+1, 'MaxmA'] = final_PROTdf.loc[row, 'MaxmA'] 
                            final_PROTdf.loc[row, 'MaxmA'] = ''
                            final_PROTdf.loc[row+1, 'Kernel'] = final_PROTdf.loc[row, 'Kernel'] 
                            final_PROTdf.loc[row, 'Kernel'] = ''
                        else:
                            pass
                else:                                                           #Create AEC IQ Ref values for other machines
                    for row in final_PROTdf.index:
                        if (pd.notna(final_PROTdf.loc[row, 'SD'])) & (pd.notna(final_PROTdf.loc[row, 'SD Thick'])):
                            final_PROTdf.loc[row, 'AEC IQ Ref'] = str(final_PROTdf.loc[row, 'SD']) + ' @ ' + str(final_PROTdf.loc[row, 'SD Thick'])
                        else:
                            pass
                final_PROTdf.drop(['SD', 'SD Thick'], axis = 1, inplace = True)     #Drop SD and SD thick rows
            
                final_PROTdf.to_excel(writer, sheet_name = machinename, index=False)
                print(machinename)
            
    finaldesname = 'Print_Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    despath = os.path.join(dirname, finaldesname)
    
    PDFformatter.PDF_formatter(finalpath, despath)

            
