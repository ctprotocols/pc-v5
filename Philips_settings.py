# -*- coding: utf-8 -*-
"""
Created on Thu Sep 17 11:59:02 2020

@author: EastmanE
"""

originalheaders = ['Protocol', 'Scan/Recon Type', 'View Angle', 'kV', 'mA', 'mAs', 'DoseRight Index', 'Absolute Min mAs', 'Absolute Max mAs', 'Rotation Time', 'Collimation', 'Pitch', 'Field of View', 'CTDIvol', 'Dose Notification Value CTDIvol', 'DLP', 'Thickness', 'Increment', 'Field Of View', 'Filter', 'iDose Level', 'Brain Area DoseRight Index', 'Liver Area DoseRight Index']
finalheaders = ['Protocol', 'Scan/Recon Type', 'View Angle', 'kV', 'mA', 'mAs','AEC IQ Ref', 'MinmA', 'MaxmA', 'Rot (s)', 'Coll', 'Pitch', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick',  'Int', 'DFOV', 'Kernel', 'IR', 'Brain DRI', 'Liver DRI']

keyword1 = 'Exam Dose Settings'
keyword2 = 'Series'


foldpath = r'Z:\CT Protocol Management\Protocol Downloads'
finalpath = r'\\csfp2\Safety\Medical_Physics\Emi\Code\Protocol_reformatter\outputphilips.xlsx'   #Path to write the formatted protocols
despath =  r'\\csfp2\Safety\Medical_Physics\Emi\Code\Protocol_reformatter\PhilipsPrint.xlsx'

