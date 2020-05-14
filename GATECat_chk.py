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
      0 -> does not exist   
      1 -> exists
      2 -> not applicable
    Missing files list dumped to 'warnList.xlsx'
"""

import os
import glob
import pandas as pd
import re
import numpy as np

def chkFun(xlsFile, folder):
    #Import data
    #fileTypes = ('*.pdf', '*.png', '*.jpg') # the tuple of file types
    xlsData = pd.read_excel(os.path.join(folder,xlsFile)) 
    allApps = os.listdir(os.path.join(folder,'allApps'))[1:] #First file is log file
    allFileInfo = pd.DataFrame()
    
    #Create struct
    for loop in np.arange(0,len(allApps)):
        print(allApps[loop])
        if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', '*GATE*'))):
            gateFileExist=1#'YES'
        else:
            gateFileExist=0#'*** NO **'
        
        if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Grad*'))):
            gradFileExist=1#'YES'
        else:
            gradFileExist=0#'*** NO **'
        
        if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Payment*'))):
            payFileExist=1#'YES'
        else:
            payFileExist=0#'*** NO **'
        
        if(re.search('Scheduled.+',xlsData['birth_category_desc'][loop]) or 
           re.search('Economic.+',xlsData['birth_category_desc'][loop]) or 
           re.search('Non.+',xlsData['birth_category_desc'][loop])):
            if(glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Community*')) or
               glob.glob(os.path.join(folder,'allApps',allApps[loop],'checklist', 'Economically*'))):
                casteFileExist=0#'YES'
            else:
                casteFileExist=0#'** Does not exist**'
        else:
            casteFileExist=2#'- NA -'
               
            allFileInfo=allFileInfo.append(
                        pd.DataFrame({
                                 'AppNo':[allApps[loop]], 
                                 'Name':[xlsData['full_name'][loop]],
                                 'Stream':[xlsData['exam_reg_no_4'][loop][0:2]], 
                                 'Email':[xlsData['email_id'][loop]],
                                 'Mobile':[int(xlsData['mobile'][loop])],
                                 'Category':[xlsData['birth_category_desc'][loop]],
                                 'GATEcard':[gateFileExist],
                                 'CatCert':[casteFileExist], 
                                 'GradCert':[gradFileExist],
                                 'PayRef':[payFileExist]
                                  }
                                  )
                                    )
    allFileInfo.reset_index(drop=True,inplace=True)    
    # Dump email addresses - missing docs
    warnList=allFileInfo[(allFileInfo['GATEcard']==0) | (allFileInfo['CatCert']==0) |
                (allFileInfo['GradCert']==0)][['AppNo','Name','Email', 'Mobile']]
    warnList.to_excel('warnList.xlsx')
    print('File written to: '+ os.path.join(os.getcwd(),'warnList.xlsx'))

if __name__== "__main__":    
    #Var Decl
    folder = 'D:/MHRD_apps'
    #Assumes all applications inside a folder 'allApps' nested inside <folder>
    xlsFile = 'allinfo.xlsx' 
    #Assumes excel sheet with all info is in <folder> and named <xlsFile>
    chkFun(xlsFile, folder)
    # Output dumped to xlsx file
