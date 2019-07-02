#AICER new content log for highlights, loops through all the AICER reports and builds a new log based on content created in the last week/month.
#last updated 08.03.19
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

try: highlightType = sys.argv[1] #taken from command line
except: highlightType = 'weekly'

#main script
print("New content report for highlights: " + highlightType)

lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\'
#reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
#aicerDir = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"
aicerDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\"
    
dirjoined = os.path.join(aicerDir, '*AllContentItemsExport_[0-9][0-9][0-9][0-9].csv') # this joins the aicerDir variable with the filenames within it, but limits it to filenames ending in 'AllContentItemsExport_xxxx.csv', note this is hardcoded as 4 digits only. This is not regex but unix shell wildcards, as far as I know there's no way to specifiy multiple unknown amounts of numbers, hence the hardcoding of 4 digits. When the aicer report goes into 5 digits this will need to be modified, should be a few years until then though
files = sorted(glob.iglob(dirjoined), key=os.path.getctime, reverse=True) #search aicerDir and add all files to dict
filename = files[0]
reportFilename = re.search('.*\\\\AICER\\\\([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_newcontent.csv"
filename = re.search('.*\\\\AICER\\\\([^\.]*\.csv)',filename).group(1)   

date =  str(time.strftime("%d/%m/%Y"))
timeago = (datetime.datetime.now().date() - datetime.timedelta(7)) #default set to weekly
if highlightType == 'weekly': timeago = (datetime.datetime.now().date() - datetime.timedelta(7)) #the 'date' part of this means it will only provide the date, not the hours, min, sec etc
if highlightType == 'monthly': timeago = (datetime.datetime.now().date() - datetime.timedelta(30))
daterange = timeago.strftime('%b %d') + ' - ' + time.strftime('%b %d, %Y')

print('Scanning the most recent AICER report: ' + filename)
#filter
df = pd.read_csv(aicerDir + filename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
df = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
df1 = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist', 'QandAs'])] # These are the content types to keep in the report

df = df[df.DateFirstPublished.notnull()]
df['DateFirstPublished'] = pd.to_datetime(df['DateFirstPublished'], dayfirst=True)

print("Grabbing docs that were created after: " + str(timeago))
df = df[df.DateFirstPublished.dt.date > timeago]

#dropping unnecessary columns
del df['DisplayId'], df['LexisSmartId'], df['OriginalContentItemPA'], df['TopicTreeLevel4'], df['TopicTreeLevel5'], df['TopicTreeLevel6'], df['TopicTreeLevel7'], df['TopicTreeLevel8'], df['TopicTreeLevel9'], df['TopicTreeLevel10'], df['VersionFilename'], df['Filename_Or_Address'], df['CreateDate'], df['MajorUpdateFirstPublished'], df['LastPublishedDate'], df['OriginalLastPublishedDate'], df['LastMajorDate'], df['LastMinorDate'], df['LastReviewedDate'], df['LastUnderReviewDate'], df['SupportsMiniSummary'], df['HasMiniSummary']
df = df.rename(columns={'id': 'DocID', 'TopicTreeLevel1': 'PA', 'TopicTreeLevel2': 'Subtopic', 'Label': 'DocTitle'})    
df['Subtopic'] = df['Subtopic'] + ' > ' + df['TopicTreeLevel3'] 
del df['TopicTreeLevel3']


#searching for shortcuts
i = 0
for index, row in df.iterrows():
    DocID = df.iloc[i,0]
    if any(df1['OriginalContentItemId'] == DocID) == True:
        ContentItemType = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'ContentItemType'].iloc[0])
        PageType = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'PageType'].iloc[0])
        PA = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel1'].iloc[0])
        Subtopic = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel2'].iloc[0]) + ' > ' + str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel3'].iloc[0])
        DocTitle = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'Label'].iloc[0])
        DateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[0])
        OriginalContentItemId = str(DocID)
        DocID = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'id'].iloc[0])
        dictionary_row = {"DocID":DocID,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished}
        df = df.append(dictionary_row, ignore_index=True)
    i=i+1

print('Total found: ' + str(len(df)))

df.to_csv(reportDir + reportFilename, sep=',',index=False, encoding='utf-8')
print('Exported to ' + reportDir + reportFilename)

print('Finished')
