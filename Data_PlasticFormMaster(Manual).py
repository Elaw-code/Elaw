#!/usr/bin/env python
# coding: utf-8

# In[6]:


def extractFileList_Reg(pathFolder, reg, pathExport,rptType,writer):
    
    # Import pandas library
    import pandas as pd
    import os
    import datetime
    import re
    

    # Create the pandas DataFrame with column name is provided explicitly
    df = pd.DataFrame(columns=['Folder','File','CreatedTime','FileName'])
    
    regPattern = re.compile(reg)

    ctFolder=0

    for root, dirs, files in os.walk(pathFolder):
        ctFolder = ctFolder +1
        print(str(ctFolder) + ":" + root)

        for file in files:
            filePath=root + "//" + file

            regResult = regPattern.match(file)

            if regResult is None:
                print(file + " - Pattern not found.")
                continue
            else:
                print(file + " - Pattern found.")

            fileDT=os.path.getmtime(filePath)
            fileDTFomrat = datetime.datetime.fromtimestamp(fileDT)
            
            filename=re.sub("[A-Za-z-._ ()]+pdf", "", file)
            filename = filename[0:12]

            dfTemp = pd.DataFrame([{'Folder':root, 'File':file,'CreatedTime':fileDTFomrat,'FileName':filename}])
            df = pd.concat([df, dfTemp], axis=0, ignore_index=True)

    df.to_excel(writer, sheet_name=rptType, index=True)
    pathExport=pathExport.replace("'","''")


# In[7]:


#============================
import shutil
import pandas as pd
from datetime import datetime

now = datetime.now()
timestamp = now.strftime("%Y%m%d%H%M%S") #yyyyMMddHHmmss
today = now.strftime("%Y%m%d")#yyyyMMdd

# https://regexr.com/
reg = "[0-9]+.BCN-(.*?)+.pdf"

pathExport = r"K:\process_improvement\10_Shared_IT\BCN Plastic Form\Master\BCN_Master.xlsx"

# create an empty workbook
writer = pd.ExcelWriter(pathExport, engine='xlsxwriter')

pathHGFolder =r"K:\CF_Logistics_Hardgoods\02_Dokumente\01_Hardgoods\13_Others\05. Plastic declaration"

pathTXFolder =r"K:\Logistics_Textile\02_Dokumente\23_E_documents\Cert. plastics duty"

#***********************************************************************************
# For HG
teams = "HG"
extractFileList_Reg(pathHGFolder,reg,pathExport,teams,writer)
# For TX
teams = "TX"
extractFileList_Reg(pathTXFolder,reg,pathExport,teams,writer)
#***********************************************************************************
writer.save()
writer.close()
#***********************************************************************************


# In[ ]:




