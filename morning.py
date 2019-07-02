#AICER updated content log for highlights, loops through all the most recent AICER report and builds a log based on content updated/new in the last week/month
#last updated 16.05.19
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

def Filter(reportDir, filename, df, dfshortcuts, highlightType, updatenewtype):    
    print(updatenewtype, highlightType)
    date =  str(time.strftime("%d%m%Y"))
    if highlightType == 'weekly': timeago = (datetime.datetime.now().date() - datetime.timedelta(8)) #the 'date' part of this means it will only provide the date, not the hours, min, sec etc
    if highlightType == 'monthly': timeago = (datetime.datetime.now().date() - datetime.timedelta(32))
    daterange = timeago.strftime('%b %d') + ' - ' + time.strftime('%b %d, %Y')

    df = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    dfshortcuts = dfshortcuts[dfshortcuts.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    
    df['LastPublishedDate'] = pd.to_datetime(df['LastPublishedDate'], dayfirst=True)
    df['LastUnderReviewDate'] = pd.to_datetime(df['LastUnderReviewDate'], dayfirst=True)

    def UnderReview(df):
        if df['LastPublishedDate'] < df['LastUnderReviewDate']:
            val = True
        else:
            val = False
        return val

    df['UnderReview'] = df.apply(UnderReview, axis =1)

    #df1 = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    
    if updatenewtype == 'updated':
        df = df[df.MajorUpdateFirstPublished.notnull()]
        df1 = df
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        
        df['MajorUpdateFirstPublished'] = pd.to_datetime(df['MajorUpdateFirstPublished'], dayfirst=True)
        print("Grabbing docs that were updated after: " + str(timeago))
        df = df[df.MajorUpdateFirstPublished.dt.date > timeago]
        reportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_updated_" + date + ".csv"
   
    if updatenewtype == 'new':
        df = df[df.DateFirstPublished.notnull()]
        df1 = df
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist', 'QandAs'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'Q&As', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        
        df['DateFirstPublished'] = pd.to_datetime(df['DateFirstPublished'], dayfirst=True)
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
                    OriginalContentItemId = str(DocID)
                    DocID2 = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'id'].iloc[x])
                    UR = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'UnderReview'].iloc[x])
                    MajorUpdateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'MajorUpdateFirstPublished'].iloc[x])
                    DateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                    dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                    
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
                            OriginalContentItemId = str(DocID)
                            DocID2 = '0'
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            MajorUpdateFirstPublished = 'nan'
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            
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
                            OriginalContentItemId = str(DocID)
                            DocID2 = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'id'].iloc[x])
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            MajorUpdateFirstPublished = 'nan'
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            
                            df = df.append(dictionary_row, ignore_index=True)                    
                            listofpas.append(PA)
                                                
        i=i+1

    print('Total found: ' + str(len(df)))
    if updatenewtype == 'updated':
        columnsTitles = ['DocID', 'ContentItemType', 'OriginalContentItemId', 'PageType', 'PA', 'Subtopic', 'DocTitle', 'MajorUpdateFirstPublished', 'UnderReview']
    
    if updatenewtype == 'new':
        columnsTitles = ['DocID', 'ContentItemType', 'OriginalContentItemId', 'PageType', 'PA', 'Subtopic', 'DocTitle', 'DateFirstPublished', 'UnderReview']
    
    df = df.reindex(columns=columnsTitles)

    df.to_csv(reportDir + reportFilename, sep=',',index=False, encoding='utf-8')
    print('Exported to ' + reportDir + reportFilename)

#main script
print("New and Updated content report for highlights...\n\n")


#reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
#aicerDir = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"
aicerDir = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"
globalmetricsDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Global_Metrics\\'
shortcutnodeaicer = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\AllContentItemsExportWithShortCutNodeInfo.csv'
print('Loading the Aicer report can take up to 20 minutes depending on your connection, please be patient...\n\n')
dirjoined = os.path.join(aicerDir, '*AllContentItemsExport_[0-9][0-9][0-9][0-9].csv') # this joins the aicerDir variable with the filenames within it, but limits it to filenames ending in 'AllContentItemsExport_xxxx.csv', note this is hardcoded as 4 digits only. This is not regex but unix shell wildcards, as far as I know there's no way to specifiy multiple unknown amounts of numbers, hence the hardcoding of 4 digits. When the aicer report goes into 5 digits this will need to be modified, should be a few years until then though
files = sorted(glob.iglob(dirjoined), key=os.path.getctime, reverse=True) #search aicerDir and add all files to dict
filename = files[0]
filename = re.search('.*\\\\AICER\\\\([^\.]*\.csv)',filename).group(1)
print('Scanning the most recent AICER report: ' + filename)
#filter
dfaicer = pd.read_csv(aicerDir + filename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
dfshortcuts =  pd.read_csv(globalmetricsDir + 'AllContentItemsExportWithShortCutNodeInfo.csv', encoding='utf-8', low_memory=False) #Load csv file into dataframe

Filter(reportDir, filename, dfaicer, dfshortcuts, 'weekly', 'new')
Filter(reportDir, filename, dfaicer, dfshortcuts, 'weekly', 'updated')
Filter(reportDir, filename, dfaicer, dfshortcuts, 'monthly', 'new')
Filter(reportDir, filename, dfaicer, dfshortcuts, 'monthly', 'updated')

print('Finished')
