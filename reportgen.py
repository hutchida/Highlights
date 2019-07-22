#AICER updated content log for highlights, loops through all the most recent AICER report and builds a log based on content updated/new in the last week/month
#last updated 17.06.19
#Developed by Daniel Hutchings

import csv
import pandas as pd
import glob
import os
import fnmatch
import re
import time
import datetime
import sys
import xml.etree.ElementTree as ET
from lxml import etree

def Filter(reportDir, filename, df, dfshortcuts, highlightType, updatenewtype):    
    print(updatenewtype, highlightType)
    date =  str(time.strftime("%d%m%Y"))
    if highlightType == 'weekly': timeago = (datetime.datetime.now().date() - datetime.timedelta(8)) #the 'date' part of this means it will only provide the date, not the hours, min, sec etc
    if highlightType == 'monthly': timeago = (datetime.datetime.now().date() - datetime.timedelta(32))
    daterange = timeago.strftime('%b %d') + ' - ' + time.strftime('%b %d, %Y')

    df = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    dfshortcuts = dfshortcuts[dfshortcuts.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    
    df['LastPublishedDate'] = pd.to_datetime(df['LastPublishedDate'], dayfirst=False) #Date is in American format hence dayfirst false
    df['LastUnderReviewDate'] = pd.to_datetime(df['LastUnderReviewDate'], dayfirst=False) #Date is in American format hence dayfirst false

    def UnderReview(df):
        if df['LastPublishedDate'] < df['LastUnderReviewDate']:
            val = True
        else:
            val = False
        return val

    df['UnderReview'] = df.apply(UnderReview, axis =1)
    df['OriginalContentItemType'] = ''
    df['OriginalContentItemPA'] = ''
    
    df['MajorUpdateFirstPublished'] = pd.to_datetime(df['MajorUpdateFirstPublished'], dayfirst=False) #US date
    df['DateFirstPublished'] = pd.to_datetime(df['DateFirstPublished'], dayfirst=False) #US date

    if updatenewtype == 'updated':
        df = df[df.MajorUpdateFirstPublished.notnull()]
        df1 = df
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        
        
        print("Grabbing docs that were updated after: " + str(timeago))
        df = df[df.MajorUpdateFirstPublished.dt.date > timeago]
        reportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_updated_" + date + ".csv"
   
    if updatenewtype == 'new':
        df = df[df.DateFirstPublished.notnull()]
        df1 = df
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist', 'QandAs'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'Q&As', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        
        
        print("Grabbing docs that were created after: " + str(timeago))
        df = df[df.DateFirstPublished.dt.date > timeago]
        reportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_new_" + date + ".csv"
   

    #dropping unnecessary columns
    del df['DisplayId'], df['LexisSmartId'], df['OriginalContentItemPA'], df['TopicTreeLevel4'], df['TopicTreeLevel5'], df['TopicTreeLevel6'], df['TopicTreeLevel7'], df['TopicTreeLevel8'], df['TopicTreeLevel9'], df['TopicTreeLevel10'], df['VersionFilename'], df['Filename_Or_Address'], df['CreateDate'], df['LastPublishedDate'], df['OriginalLastPublishedDate'], df['LastMajorDate'], df['LastMinorDate'], df['LastReviewedDate'], df['LastUnderReviewDate'], df['SupportsMiniSummary'], df['HasMiniSummary']
    df = df.rename(columns={'id': 'DocID', 'TopicTreeLevel1': 'PA', 'TopicTreeLevel2': 'Subtopic', 'Label': 'DocTitle'})    
    df['Subtopic'] = df['Subtopic'] + ' > ' + df['TopicTreeLevel3'] 
    del df['TopicTreeLevel3']


    #searching for shortcuts
    print('Searching for shortcuts...')
    i = 0
    #df.to_csv(reportDir + 'test-df-' + updatenewtype + '-' + highlightType + '.csv', sep=',',index=False, encoding='utf-8')

    for index, row in df.iterrows():
        DocID = df.iloc[i,0]
        masterPA = df.iloc[i,4]
        masterContentType = df.iloc[i,1]
        listofpas = [masterPA]
        
        #individual shortcuts
        if any(df1['OriginalContentItemId'] == DocID) == True:
            shortcutcount = len(df1[df1.OriginalContentItemId.isin([DocID])]) #filter by OCI with DocID to find how many shortcuts
            for x in range(0,shortcutcount):       
                PA = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel1'].iloc[x])
     
                if PA not in listofpas:    
                    
                    ContentItemType = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'ContentItemType'].iloc[x])
                    PageType = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'PageType'].iloc[x])
                    Subtopic = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel2'].iloc[x]) + ' > ' + str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel3'].iloc[x])
                    DocTitle = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'Label'].iloc[x])
                    
                    OriginalContentItemId = int(DocID)
                                   
                    DocID2 = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'id'].iloc[x])
                    UR = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'UnderReview'].iloc[x])
                    MajorUpdateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'MajorUpdateFirstPublished'].iloc[x])
                    DateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                    dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemType":masterContentType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA, "OriginalContentItemPA":masterPA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            
                    df = df.append(dictionary_row, ignore_index=True)                    
                    listofpas.append(PA)
        
        if any(dfshortcuts['id'] == DocID) == True:
            doccount = len(dfshortcuts[dfshortcuts.id.isin([DocID])]) #filter by id with DocID to find how many shortcuts
            if doccount > 1 :                
                for x in range(1,doccount):    
                    shortcutsubtopic = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'IsSubTopicShortcut'].iloc[x])
                    if shortcutsubtopic == 'Y':                       
                        PA = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'TopicTreeLevel1'].iloc[x])        
                        print(x, doccount, PA, DocID)    
                        if PA not in listofpas:   
                            ContentItemType = 'SubtopicShortcut'
                            PageType = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'PageType'].iloc[x])
                            Subtopic = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'TopicTreeLevel2'].iloc[x]) + ' > ' + str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'TopicTreeLevel3'].iloc[x])
                            DocTitle = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'Label'].iloc[x])
                            OriginalContentItemId = int(DocID)
                            DocID2 = '0'
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            MajorUpdateFirstPublished = 'nan'
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemType":masterContentType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA, "OriginalContentItemPA":masterPA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            df = df.append(dictionary_row, ignore_index=True)                    
                            listofpas.append(PA)
                            
        if any(dfshortcuts['OriginalContentItemId'] == DocID) == True:
            doccount = len(dfshortcuts[dfshortcuts.OriginalContentItemId.isin([DocID])]) #filter by id with DocID to find how many shortcuts
            if doccount > 1 :                
                for x in range(1,doccount):    
                    shortcutsubtopic = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'IsSubTopicShortcut'].iloc[x])
                    if shortcutsubtopic == 'Y':                       
                        PA = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'TopicTreeLevel1'].iloc[x])        
                        print(x, doccount, PA, DocID)    
                        if PA not in listofpas:   
                            ContentItemType = 'SubtopicShortcutOfShortcut'
                            PageType = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'PageType'].iloc[x])
                            Subtopic = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'TopicTreeLevel2'].iloc[x]) + ' > ' + str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'TopicTreeLevel3'].iloc[x])
                            DocTitle = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'Label'].iloc[x])
                            OriginalContentItemId = int(DocID)
                            DocID2 = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'id'].iloc[x])
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            MajorUpdateFirstPublished = 'nan'
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemType":masterContentType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA, "OriginalContentItemPA":masterPA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            df = df.append(dictionary_row, ignore_index=True)                    
                            listofpas.append(PA)
                                                
        i=i+1

    print('Total found: ' + str(len(df)))
    #cleanup
    df.OriginalContentItemId = df.OriginalContentItemId.fillna(0) #fill all empty values of this column with zeros
    df.OriginalContentItemId = df.OriginalContentItemId.astype(int) #convert hidden float column to int to remove trailing decimals when exporting to csv

    if updatenewtype == 'updated':
        columnsTitles = ['DocID', 'OriginalContentItemId', 'ContentItemType', 'OriginalContentItemType', 'PageType', 'PA', 'OriginalContentItemPA', 'Subtopic', 'DocTitle', 'MajorUpdateFirstPublished', 'UnderReview']
    
    if updatenewtype == 'new':
        columnsTitles = ['DocID', 'OriginalContentItemId', 'ContentItemType', 'OriginalContentItemType', 'PageType', 'PA', 'OriginalContentItemPA', 'Subtopic', 'DocTitle', 'DateFirstPublished', 'UnderReview']
    
    df = df.reindex(columns=columnsTitles)

    df.to_csv(reportDir + reportFilename, sep=',',index=False, encoding='utf-8')
    print('Exported to ' + reportDir + reportFilename)
    return(reportDir + reportFilename)



def FindMostRecentFile(directory, pattern):
    filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
    filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
    return filelist[0]



#main script
env = sys.argv[1] #taken from command line
print("New and Updated content report for highlights...")

#Directories
if env == 'dev': 
    reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
    print('Export directory set to DEV folder...')
else: reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'

#reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
aicerDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER\\'
globalmetricsDir = '\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER_withShortcut_AdHoc\\'
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'

aicerFilename = FindMostRecentFile(aicerDir, '*AICER*.csv')
aicerFilename = re.search('.*\\\\AICER\\\\([^\.]*\.csv)',aicerFilename).group(1)
print('Loading the most recent AICER report: ' + aicerFilename)
aicershortcutsFilename = FindMostRecentFile(globalmetricsDir, 'AllContentItemsExportWithShortCutNodeInfo_*.csv')
print('Loading the most recent AICER Shortcuts report: ' + aicershortcutsFilename)
#filter
dfaicer = pd.read_csv(aicerDir + aicerFilename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
print('Aicer loaded...loading Aicer shortcuts...')
dfshortcuts =  pd.read_csv(aicershortcutsFilename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
print('Aicer shortcuts loaded...filtering reports...')

Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'weekly', 'new')
Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'weekly', 'updated')
Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'monthly', 'new')
Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'monthly', 'updated')
