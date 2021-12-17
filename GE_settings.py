# -*- coding: utf-8 -*-
"""
Created on Thu Sep 17 11:59:02 2020

@author: EastmanE
"""

# originalheaders = ['Protocol', 'Scan/Recon Type', 'AutoStore', 'Gating', ' SeriesLevelCopy', 'Injector', 'Scan', 'kV', 'mA', 'Plane', 'Message', 'Timer', 'Light', 'End', 'SmartPrep', 'Speed', 'Type', 'Rows', 'Int', 'HiRes', 'Shuttle', 'Tilt', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'DLP NV', 'Images', 'Thick', 'DFOV', 'A/P', 'R/L', 'Filter', 'ASIR', 'Flip', 'SmartmA', 'NoiseIndex', 'MinmA', 'MaxmA', 'IQEnhance', 'Smart Prep Parameters', 'pitch']
# genericheaders = ['Protocol', 'Scan/Recon Type', 'AutoStore', 'Gating', 'SeriesLevelCopy', 'Injector', 'Scan', 'kV', 'mA', 'Plane', 'Message', 'Timer', 'Light', 'End', 'SmartPrep', 'Rot (s)', 'Type', 'Coll', 'Int', 'HiRes', 'Shuttle', 'Tilt', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'DLP NV', 'Images', 'Thick', 'DFOV', 'A/P', 'R/L', 'Kernel', 'IR', 'Flip', 'SmartmA', 'IQ Ref', 'MinmA', 'MaxmA', 'IQEnhance', 'Smart Prep Parameters', 'Pitch']
# genericheaders = ['Protocol', 'Scan/Recon Type', 'Scan', 'kV', 'mA', 'Plane', 'SmartPrep', 'Rot (s)', 'Type', 'Coll', 'Int', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'DLP NV', 'Thick', 'DFOV', 'Kernel', 'IR', 'IQ Ref', 'MinmA', 'MaxmA', 'Pitch', 'IQEnhance', 'HiRes']

originalheaders = ['Protocol', 'Scan/Recon Type', 'Type', 'Plane', 'kV', 'mA', 'NoiseIndex', 'MinmA', 'MaxmA', 'Speed', 'Rows', 'pitch', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick', 'Int', 'DFOV', 'Filter', 'ASIR', 'IQEnhance', 'HiRes', 'Mode', 'Group', 'ECGMinmA', 'ECGMaxmA', 'GSI']
finalheaders = ['Protocol', 'Scan/Recon Type', 'Type', 'Plane','kV', 'mA', 'AEC IQ Ref', 'MinmA', 'MaxmA', 'Rot (s)', 'Coll', 'Pitch', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick',  'Int','DFOV', 'Kernel', 'IR', 'IQEnhance', 'HiRes', 'Mode', 'Group']

keyword1 = 'Exam Dose Settings'
keyword2 = 'Series'



# finalfields = ['Protocol', 'Scan Type', ' kV', ' mA', ' Plane', ' SmartPrep', 'Speed', ' Rows', ' Int', ' SFOV', ' CTDI', ' CTDI NV', ' DLP', ' DLP NV', ' Thick', ' DFOV', ' Filter', ' ASIR', ' NoiseIndex', ' MinmA', ' MaxmA', ' pitch', ' IQEnhance', ' HiRes']
# renamefields = ['Protocol', 'Scan', 'kV', 'mA', 'SD', 'SD Thick', 'MinmA', 'MaxmA', 'Rot (s)', 'Acq', 'Pitch', 'CFOV', 'CTDI', 'DLP', 'ReconProcess', 'Thick (mm)', 'DFOV (cm)', 'FC', 'SEMAR', 'SUREIQ']
