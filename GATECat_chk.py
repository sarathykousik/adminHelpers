# -*- coding: utf-8 -*-
"""
Created on Wed May 13 20:20:08 2020

@author: kousik
Functions:
    1.Check if GATE score card exists
    2.Check if community certificate exists
    3.Create a summary list -- allFileInfo
        | App no| Name | Category | GATE score exists? | Community certificate exists?|
    Next up:
         + open file + wait till file is closed + continue
         
"""

import os
import glob
import pandas as pd
import re
import numpy as np

#Var Decl
folder = 'D:/MHRD_apps'
#Assumes all applications inside a folder 'allApps' nested inside <folder>
xlsFile = 'allinfo.xlsx' 
#Assumes excel sheet with all info is in <folder> and named <xlsFile>
###################################
# Output dumped to 'allFileInfo'
###################################

#Import data
#fileTypes = ('*.pdf', '*.png', '*.jpg') # the tuple of file types
xlsData = pd.read_excel(os.path.join(folder,xlsFile)) 
allApps = os.listdir(os.path.join(folder,'allApps'))[1:] #First file is log file
allFileInfo = pd.DataFrame([])

#Create struct
for loop in np.arange(0,len(allApps)):
    print(allApps[loop])
    if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', '*GATE*'))):
        gateFileExist='YES'
    else:
        gateFileExist='*** NO **'
    
    if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Grad*'))):
        gradFileExist='YES'
    else:
        gradFileExist='*** NO **'
    
    if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Payment*'))):
        payFileExist='YES'
    else:
        payFileExist='*** NO **'
    
    if(re.search('Scheduled.+',xlsData['birth_category_desc'][loop]) or 
       re.search('Economic.+',xlsData['birth_category_desc'][loop]) or 
       re.search('Non.+',xlsData['birth_category_desc'][loop])):
        if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Community*')) or
           glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Economically*'))):
            casteFileExist='Exists'
        else:
            casteFileExist='** Does not exist**'
    else:
        casteFileExist='- NA -'
           
    allFileInfo = allFileInfo.append(
                    pd.DataFrame([allApps[loop], xlsData['full_name'][loop], 
                         xlsData['birth_category_desc'][loop],gateFileExist,    
                         casteFileExist, gradFileExist,payFileExist]).transpose()
                                     )
allFileInfo.columns=[['App No', 'Name', 'Category', 'GATEcard?', 
                                  'CatCert?', 'GradCert?','PayRef?' ]]
        
    