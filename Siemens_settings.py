# -*- coding: utf-8 -*-
"""
Created on Mon Sep 21 10:05:05 2020

@author: EastmanE
"""

columnheaders = [' Pitch2', 'Region/Protocol Name', 'Range', 'Voltage', 'MA', 'QualityRefMAs', 'AECReferenceMAs', 'CustomMAs', 'PitchFactor', 'MinmA', 'MaxmA', 'RotTime', 'Acq.', 'SFOV', 'CTDIw', 'DoseNotificationValueCTDIvol', 'DLP', 'ReconSliceEffective', 'SliceEffective', 'ReconIncr', 'DFOV', 'Kernel', 'SAFIREStrength', 'CustomMAsA', 'CustomMAsB', 'SeriesDescription', 'CAREkV', 'RefKV']
finalheaders = ['Pitch2', 'Region/Protocol Name', 'Scan/Recon Type', 'kV', 'mA', 'QualityRefMAs', 'AECReferenceMAs', 'CustomMAs', 'Pitch', 'MinmA', 'MaxmA', 'Rot (s)', 'Coll', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick', 'SliceEffective', 'Int', 'DFOV', 'Kernel', 'IR', 'CustomMAsA', 'CustomMAsB', 'SeriesDescription', 'CAREkV', 'RefKV']

renamefields = ['Protocol', 'Scan/Recon Type', 'kV', 'mA', 'AEC IQ Ref', 'Pitch', 'MinmA', 'MaxmA', 'Rot (s)', 'Coll', 'SFOV', 'CTDI', 'CTDI NV', 'DLP', 'Thick', 'Int', 'DFOV', 'Kernel',  'IR', 'CAREkV', 'RefKV']
