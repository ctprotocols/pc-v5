# -*- coding: utf-8 -*-
"""
Created on Thu Sep 17 11:59:02 2020

@author: EastmanE
"""

# fields = ['Protocol Name','kV', 'mA', 'MinmA', 'MaxmA', 'SliceThicknessRecon', 'FC', 'ReconProcess', 'ScanModeName', 'ScanoAngle', 'CFOV', 'CTDI', 'Collimation', 'DLP', 'Pitch', 'RotationTime', 'Reconstruction', 'DFOV', 'SEMAR', 'SUREIQ', 'SD lookup','SD', 'SliceThickness', 'XYModulation']
 
fields = ['ExamPlanName', 'ExamPlanNo', 'SliceThickness', 'SliceInterval', 'FC', 'ReconProcess', 'ScanModeParam', 'ScanModeName', 'kV', 'mA', 'CFOV', 'CTDI', 'Collimation', 'DLP', 'Pitch', 'RotationTime', 'Reconstruction', 'ReconProcess', 'DFOV', 'SEMAR', 'SUREIQ', 'SureExpOrgan', 'SureExpName']
MOTfields = ['ExamPlanName', 'ExamPlanNo', 'ScanoModeName', 'ScanokV', 'ScanomA', 'FC', 'ScanModeParam', 'ScanModeName', 'Collimation', 'Pitch', 'kV', 'mA', 'RotationTime', 'CFOV', 'CTDI', 'DLP', 'Reconstruction', 'SliceThickness', 'ReconProcess', 'SliceInterval', 'DFOV', 'SXTargetSD', 'SXTargetSliceThickness', 'SXFC', 'SXMaxMA', 'SXMinMA', 'SureIQRecon']


MOTfinalfields2 = ['Protocol Name', 'ScanModeName', 'kV', 'mA', 'SXTargetSD', 'SXTargetSliceThickness', 'SXMinMA', 'SXMaxMA', 'RotationTime', 'Collimation', 'Pitch', 'CFOV', 'CTDI', 'CTDI NV', 'DLP',  'SliceThickness', 'SliceInterval', 'DFOV', 'FC', 'ReconProcess', 'SureIQRecon', 'SEMAR']
finalfields2 = ['Protocol Name', 'ScanModeName', 'kV', 'mA', 'SD', 'SD Thick', 'MinmA', 'MaxmA', 'RotationTime', 'Collimation', 'Pitch', 'CFOV', 'CTDI', 'CTDI NV', 'DLP', 'SliceThickness', 'SliceInterval', 'DFOV', 'FC', 'ReconProcess', 'SUREIQ',  'SEMAR']

renamefields = ['Protocol', 'Scan/Recon Type', 'kV', 'mA', 'SD', 'SD Thick', 'MinmA', 'MaxmA', 'Rot (s)', 'Coll', 'Pitch', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick', 'Int', 'DFOV', 'Kernel',  'IR',  'SUREIQ', 'SEMAR']

reconfields = ['SureIQ Name', 'FC', 'AIDR', 'Filter', 'Boost3D', 'OSR', 'WL1', 'WL2', 'WL3', 'WW1', 'WW2', 'WW3']


# originalheaders = ['Protocol', 'Scan/Recon Type', 'Type', 'SmartPrep', 'Plane', 'kV', 'mA', 'NoiseIndex', 'MinmA', 'MaxmA', 'Speed', 'Rows', 'pitch', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick', 'Int', 'DFOV', 'Filter', 'ASIR', 'IQEnhance', 'HiRes']
# finalheaders = ['Protocol', 'Scan/Recon Type', 'Type', 'SmartPrep', 'Plane','kV', 'mA', 'AEC IQ Ref', 'MinmA', 'MaxmA', 'Rot (s)', 'Coll', 'Pitch', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick',  'Int','DFOV', 'Kernel', 'IR', 'IQEnhance', 'HiRes']

# keyword1 = 'Exam Dose Settings'
# keyword2 = 'Series'

