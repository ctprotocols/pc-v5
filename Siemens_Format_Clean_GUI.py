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
import xml.etree.ElementTree as ET
from Siemens_settings import columnheaders, finalheaders, renamefields
import PDFformatter

def Siemens(filename, machinename, typeofscanner):
    dirname = os.path.dirname(filename)
    finalfilename = 'Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    
    finalpath = os.path.join(dirname, finalfilename)
    
    year = dt.datetime.today().strftime('%Y')
    
    siempaths = [filename]
    # =============================================================================
    # Read files (must be xlsx file)
    # =============================================================================
    
    with pd.ExcelWriter(finalpath) as writer:
        for exportfile in siempaths:
    
            tree = ET.parse(exportfile)                                         #Parse the xml file
            root = tree.getroot()                                                   
            
            taglist = [elem.tag for elem in root.iter()]                        #Write the keywords to a list
            textlist = [elem.text for elem in root.iter()]                      #Write the values to a list
            
            df = pd.DataFrame()                                                 #Write the lists to a df for
            df[0] = taglist                                                     #easier analysis
            df[1] = textlist
            
            
            folderidx = df[df[0] == 'FolderName'].index.tolist()    #Start of custom protocols and default protocols
            dfend = df.index[-1]                                     #Last line in dataframe
            folderidx.append(dfend)
            # =============================================================================
            # Break down the data by sections
            # =============================================================================
            
            if 'Siemens Intevo BOLD' in typeofscanner:                            #Sliceval tells the code where to stop looking
                sliceval = -1                                   
            else:
                sliceval = -2
            for idx,val in enumerate(folderidx[:sliceval]):               #Break up the data into sections    
                compiledf = pd.DataFrame(columns = columnheaders)           #Empty df to collect reorganized data
                protocoltype = df.loc[val, 1]                               #Break into first sections based on protocol
                print(protocoltype)                                         #type. (Custom or default)
                start = val
                stop = folderidx[idx+1]
                df1 = df.loc[start:stop, :]                                 #Only look at one set of protocols (custom or default)
            
                bodyidx = df1[df1[0] == 'BodySize'].index.tolist()          #Find where the protocols are broken up into adult and child
                df1end = df1.index[-1]
                bodyidx.append(df1end)
                
                for idx2, val2 in enumerate(bodyidx[:-1]):          #Break up the data into sections
                    bodyvalue = df1.loc[val2, 1]                    #Adult or child
                    print(bodyvalue)
                    start2 = val2
                    stop2 = bodyidx[idx2+1]
                    df2 = df1.loc[start2:stop2, :]                  #Only look at one section (Adult or Child)
                    
                    regionidx = df2[df2[0] == 'RegionName'].index.tolist()      #Pull the anatomy region name so 
                    df2end = df2.index[-1]                                      #we can add to protocol names to make
                    regionidx.append(df2end)                                    #them unique.
                    
                    for idx3, val3 in enumerate(regionidx[:-1]):
                        regionvalue = df2.loc[val3,1]                           #Break into sections based on region
                        # print(regionvalue)
                        
                        start3 = val3                                           #Only look at one region at a time      
                        stop3 = regionidx[idx3+1]
                        df3 = df2.loc[start3:stop3, :]
                        
                        protocolidx = df3[df3[0] == 'ProtocolName'].index.tolist()  #Break up into sections 
                        df3end = df3.index[-1]                                      #based on protocol.
                        protocolidx.append(df3end)
                        
                        for idx4, val4 in enumerate(protocolidx[:-1]):
                            nameprotocol = df3.loc[val4, 1]
                            start4 = val4
                            stop4 = protocolidx[idx4+1]
                            df4 = df3.loc[start4:stop4, :]
                            
                            protocolname = df4.loc[start4, 1]                       #Protocols are broken up into scan and recon entries. break up into sections again.
                            scanidx = df4[(df4[0] == 'ScanEntry') | (df4[0] == 'ReconJob')].index.tolist()
                            df4end = df4.index[-1]
                            scanidx.append(df4end)
                            
                            for idx5, val5 in enumerate(scanidx[:-1]):      #Now sections are completely broken up to the smallest form. 
                                start5 = val5+1
                                stop5 = scanidx[idx5+1]
                                df5 = df4.loc[start5:stop5, :]
                                scantags = df5[0].tolist()                  #Column headers in this section
                                scantext = [df5[1].tolist()]                #Values in this section
                                linedf = pd.DataFrame(data = scantext, columns = scantags)  #out the info into a dataframe format
                                linedf.insert(0, 'Region/Protocol Name', str(regionvalue) + r'/' + str(nameprotocol)) #Insert the name of the protocol
                                compiledf = pd.concat([compiledf, linedf], ignore_index = True)         #Append the row to the collection df
                # compiledf.insert[0, 'Pitch2', '']
    
                compiledf = compiledf[columnheaders]    #Only keep columns that we want
                compiledf.columns = finalheaders        #Give columns generic names
                
                
                
                for row in compiledf.index:
                    if ('Fl' in str(compiledf.loc[row, 'Scan/Recon Type'])) & (pd.notna(compiledf.loc[row, 'Pitch'])):  #Divide pitch by 2 for flash protocols.
                        compiledf.loc[row, 'Pitch2'] = str(float(compiledf.loc[row, 'Pitch'])/2)
                    else:
                        compiledf.loc[row, 'Pitch2'] = compiledf.loc[row, 'Pitch']
                    if pd.notna(compiledf.loc[row, 'SliceEffective']):                  #Replace Thickness with SliceEffective if there is a value available 
                        compiledf.loc[row, 'Thick'] = compiledf.loc[row, 'SliceEffective']
                    else:
                        pass
                    if pd.isnull(compiledf.loc[row, 'Scan/Recon Type']):                #Add in the series decription to the Scan/Recon Type
                        if pd.notna(compiledf.loc[row, 'SeriesDescription']):
                            if '/' in str(compiledf.loc[row, 'SeriesDescription']):
                                compiledf.loc[row, 'SeriesDescription'] = str(compiledf.loc[row, 'SeriesDescription']).replace('/', '_')
                            else:
                                pass
                            compiledf.loc[row, 'Scan/Recon Type'] = 'Recon ' + str(compiledf.loc[row, 'SeriesDescription'])
                        else:
                            compiledf.loc[row, 'Scan/Recon Type'] = 'Recon '
    
                    else:
                        pass
                    if (pd.notna(compiledf.loc[row, 'AECReferenceMAs'])) & (pd.isna(compiledf.loc[row, 'QualityRefMAs'])):  #Same parameter with a different name. Move values to QualityRefMAs column
                        compiledf.loc[row, 'QualityRefMAs'] = compiledf.loc[row, 'AECReferenceMAs']                         #for machines that use this nomenclature
                    else:
                        pass
                    if compiledf.loc[row, 'CAREkV'] == 'Off':       #Remove the RefkV value if CarekV is Off
                        compiledf.loc[row, 'RefKV'] = ''
                    else:
                        pass
                    # =============================================================================
                    # Calculate mA from CustomMAs values (*pitch/RotationTime). Account for single and dual tubes.                
                    # =============================================================================
                    if pd.isna(compiledf.loc[row, 'mA']):
                        if (pd.notna(compiledf.loc[row, 'CustomMAs'])) & (pd.notna(compiledf.loc[row, 'Pitch2'])):
                            if float(compiledf.loc[row, 'Pitch2']) != 0:
                                compiledf.loc[row, 'mA'] = str(int(float(compiledf.loc[row, 'CustomMAs'])*float(compiledf.loc[row, 'Pitch2'])/float(compiledf.loc[row, 'Rot (s)'])))
                            else:
                                compiledf.loc[row, 'mA'] = str(int(float(compiledf.loc[row, 'CustomMAs'])/float(compiledf.loc[row, 'Rot (s)'])))
                        if (pd.notna(compiledf.loc[row,'CustomMAsA'])) & (pd.notna(compiledf.loc[row,'CustomMAsB'])):
                            if float(compiledf.loc[row, 'Pitch2']) != 0:
                                tubeA =  float(compiledf.loc[row, 'CustomMAsA'])*float(compiledf.loc[row, 'Pitch2'])/float(compiledf.loc[row, 'Rot (s)'])
                                tubeB =  float(compiledf.loc[row, 'CustomMAsB'])*float(compiledf.loc[row, 'Pitch2'])/float(compiledf.loc[row, 'Rot (s)'])
                                compiledf.loc[row, 'mA'] = str(int(tubeA)) + '(A)/' +str(int(tubeB)) + '(B)'
                            else:
                                tubeA =  float(compiledf.loc[row, 'CustomMAsA'])/float(compiledf.loc[row, 'Rot (s)'])
                                tubeB =  float(compiledf.loc[row, 'CustomMAsB'])/float(compiledf.loc[row, 'Rot (s)'])
                                compiledf.loc[row, 'mA'] = str(int(tubeA)) + '(A)/' +str(int(tubeB)) + '(B)'
                        if (pd.notna(compiledf.loc[row,'CustomMAsA'])) & (pd.isna(compiledf.loc[row,'CustomMAsB'])):
                            if float(compiledf.loc[row, 'Pitch2']) != 0:
                                tubeA =  float(compiledf.loc[row, 'CustomMAsA'])*float(compiledf.loc[row, 'Pitch2'])/float(compiledf.loc[row, 'Rot (s)'])
                                compiledf.loc[row, 'mA'] = str(int(tubeA))
                            else:
                                tubeA =  float(compiledf.loc[row, 'CustomMAsA'])/float(compiledf.loc[row, 'Rot (s)'])
                                compiledf.loc[row, 'mA'] = str(int(tubeA))
                        if (pd.isna(compiledf.loc[row,'CustomMAsA'])) & (pd.notna(compiledf.loc[row,'CustomMAsB'])):
                            if float(compiledf.loc[row, 'Pitch2']) != 0:
                                tubeB =  float(compiledf.loc[row, 'CustomMAsB'])*float(compiledf.loc[row, 'Pitch2'])/float(compiledf.loc[row, 'Rot (s)'])
                                compiledf.loc[row, 'mA'] = str(int(tubeB))
                            else:
                                tubeB =  float(compiledf.loc[row, 'CustomMAsB'])/float(compiledf.loc[row, 'Rot (s)'])
                                compiledf.loc[row, 'mA'] = str(int(tubeB))
                        else:
                            pass
                    else:
                        pass
                if pd.isna(compiledf['Int']).all():         #If the Int column is completely blank, replace the int values with the Thick values
                    compiledf['Int'] = compiledf['Thick']
                else:
                    pass
                dropindex = []                      #Empty list to collect index of rows that should be dropped. 
                for row in compiledf.index[1:]:     #Combine recon rows with the corresponding scan row above it.
                    if ('Recon' in compiledf.loc[row, 'Scan/Recon Type']) & ('Recon' not in compiledf.loc[row-1, 'Scan/Recon Type']):
                        compiledf.loc[row-1, 'Scan/Recon Type'] = str(compiledf.loc[row-1, 'Scan/Recon Type']) + '/' +str(compiledf.loc[row, 'Scan/Recon Type'])
                        for col in compiledf.columns:
                            if (pd.isnull(compiledf.loc[row-1, col])) & (pd.notna(compiledf.loc[row, col])):
                                compiledf.loc[row-1, col] = compiledf.loc[row, col]
                            else:
                                pass
                        dropindex.append(row)       #After moving parameters into the row above, we need to drop the OG row.
                    else:
                        pass
                compiledf.drop(dropindex, axis = 0, inplace = True)     #drop rows and drop unnecessary columns (used for calculations)
                compiledf.drop(['SeriesDescription', 'Pitch2', 'AECReferenceMAs', 'CustomMAs', 'CustomMAsA', 'CustomMAsB', 'SliceEffective'], axis = 1, inplace = True)
    
                compiledf.columns = renamefields        #rename columns
               
                compiledf.to_excel(writer, sheet_name = machinename, index=False)       #Write to excel
    
    finaldesname = 'Print_Formatted_Protocols_' + dt.datetime.now().strftime("%Y%m%d%H%M%S") +'.xlsx'
    despath = os.path.join(dirname, finaldesname)
    
    PDFformatter.PDF_formatter(finalpath, despath)
    
    
