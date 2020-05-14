# -*- coding: utf-8 -*-
"""
Created on Thu May 14 15:45:55 2020

@author: kousik
"""

import os
import pandas as pd
import numpy as np

#Var Decl
folder = 'D:/MHRD_apps'
#Assumes all applications inside a folder 'allApps' nested inside <folder>
xlsFile = 'allinfo.xlsx' 
xlsLocFile = 'full_df.xlsx' 

xlsData = pd.read_excel(os.path.join(folder,xlsFile)) 
xlsLocData = pd.read_excel(os.path.join(folder,xlsLocFile)) 

okayFlag=np.zeros(np.shape(xlsLocData)[0])
xlsLocData = xlsLocData.assign(GATEokay=okayFlag) 

for loop in np.arange(0,np.shape(xlsLocData)[0]):
    print(xlsLocData['Name'][loop]+' - '+
              xlsLocData['Category'][loop] +' - ' +
              xlsLocData['GATEid'][loop])
    os.startfile(xlsLocData['gateFilePath'][loop])
    okayFlag[loop]=input('----- GATE ID okay?: ')
    xlsLocData['GATEokay'][loop]=okayFlag[loop]
    
