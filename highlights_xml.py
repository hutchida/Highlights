    #AICER updated content log for highlights, loops through all the AICER reports and builds a log based on content updated in the last week/month
#last updated 15.05.19
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
    
    #timeago = (datetime.datetime.now().date() - datetime.timedelta(7)) #default set to weekly
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

    df1 = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    #df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist'])] # These are the content types to keep in the report
    
    if updatenewtype == 'updated':
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        df = df[df.MajorUpdateFirstPublished.notnull()]
        df['MajorUpdateFirstPublished'] = pd.to_datetime(df['MajorUpdateFirstPublished'], dayfirst=True)
        print("Grabbing docs that were updated after: " + str(timeago))
        df = df[df.MajorUpdateFirstPublished.dt.date > timeago]
        weeklyUpdateReportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_updated_" + date + ".csv"
   
    if updatenewtype == 'new':
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist', 'QandAs'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'Q&As', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        df = df[df.DateFirstPublished.notnull()]
        df['DateFirstPublished'] = pd.to_datetime(df['DateFirstPublished'], dayfirst=True)
        print("Grabbing docs that were created after: " + str(timeago))
        df = df[df.DateFirstPublished.dt.date > timeago]
        weeklyUpdateReportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_new_" + date + ".csv"
   

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
     
                #print(x, shortcutcount, PA)    
                #if PA == masterPA: continue
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
                            #DocID2 = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'id'].iloc[x])
                            DocID2 = '0'
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            #UR = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'UnderReview'].iloc[x])
                            MajorUpdateFirstPublished = 'nan'
                            #MajorUpdateFirstPublished = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'MajorUpdateFirstPublished'].iloc[x])
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            
                            df = df.append(dictionary_row, ignore_index=True)                    
                            listofpas.append(PA)
                            #if DocID == 3229228: 
                            #    print(x, shortcutcount, dictionary_row, listofpas)   
                            #    wait = input("Review the DF before continuing...then press enter")
                    
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
                            #DocID2 = '0'
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            #UR = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'UnderReview'].iloc[x])
                            MajorUpdateFirstPublished = 'nan'
                            #MajorUpdateFirstPublished = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'MajorUpdateFirstPublished'].iloc[x])
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            
                            df = df.append(dictionary_row, ignore_index=True)                    
                            listofpas.append(PA)
                            #if DocID == 3229228: 
                            #    print(x, shortcutcount, dictionary_row, listofpas)   
                            #    wait = input("Review the DF before continuing...then press enter")
                    
                    
        i=i+1

    print('Total found: ' + str(len(df)))
    if updatenewtype == 'updated':
        columnsTitles = ['DocID', 'ContentItemType', 'OriginalContentItemId', 'PageType', 'PA', 'Subtopic', 'DocTitle', 'MajorUpdateFirstPublished', 'UnderReview']
    
    if updatenewtype == 'new':
        columnsTitles = ['DocID', 'ContentItemType', 'OriginalContentItemId', 'PageType', 'PA', 'Subtopic', 'DocTitle', 'DateFirstPublished', 'UnderReview']
    
    df = df.reindex(columns=columnsTitles)

    df.to_csv(reportDir + weeklyUpdateReportFilename, sep=',',index=False, encoding='utf-8')
    print('Exported to ' + reportDir + weeklyUpdateReportFilename)

#main script
print("New and Updated content report for highlights...\n\n")


#reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
aicerDir = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"
globalmetricsDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Global_Metrics\\'
#aicerDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\"
shortcutnodeaicer = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\AllContentItemsExportWithShortCutNodeInfo.csv'
#shortcutnodeaicer = '\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER_withShortcut_AdHoc\\AllContentItemsExportWithShortCutNodeInfo.csv'
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'

#print('Loading the Aicer report can take up to 20 minutes depending on your connection, please be patient...\n\n')
#dirjoined = os.path.join(aicerDir, '*AllContentItemsExport_[0-9][0-9][0-9][0-9].csv') # this joins the aicerDir variable with the filenames within it, but limits it to filenames ending in 'AllContentItemsExport_xxxx.csv', note this is hardcoded as 4 digits only. This is not regex but unix shell wildcards, as far as I know there's no way to specifiy multiple unknown amounts of numbers, hence the hardcoding of 4 digits. When the aicer report goes into 5 digits this will need to be modified, should be a few years until then though
#files = sorted(glob.iglob(dirjoined), key=os.path.getctime, reverse=True) #search aicerDir and add all files to dict
#filename = files[0]
#filename = re.search('.*\\\\AICER\\\\([^\.]*\.csv)',filename).group(1)
#print('Scanning the most recent AICER report: ' + filename)
#filter
#dfaicer = pd.read_csv(aicerDir + filename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
#dfshortcuts =  pd.read_csv(globalmetricsDir + 'AllContentItemsExportWithShortCutNodeInfo.csv', encoding='utf-8', low_memory=False) #Load csv file into dataframe

#Filter(reportDir, filename, dfaicer, dfshortcuts, 'weekly', 'new')
#Filter(reportDir, filename, dfaicer, dfshortcuts, 'weekly', 'updated')
#Filter(reportDir, filename, dfaicer, dfshortcuts, 'monthly', 'new')
#Filter(reportDir, filename, dfaicer, dfshortcuts, 'monthly', 'updated')

#Construct(reportDir)


def DFCleanup(df, ShortcutTypeList, ReportType):
    df1 = pd.DataFrame()
    i=0
    for index, row in df.iterrows(): #loop through df deconstruct/reconstruct
        DocID = df.DocID.iloc[i]
        ContentItemType = df.ContentItemType.iloc[i]
        PA = df.PA.iloc[i]
        OriginalContentItemPA = PA
        DocTitle = df.DocTitle.iloc[i]
        if ShortcutTypeList.count(ContentItemType) > 0: #if found in list
            ContentItemType = df.OriginalContentItemType.iloc[i]
            DocID = df.OriginalContentItemId.iloc[i]
            OriginalContentItemPA = df.OriginalContentItemPA.iloc[i]

        list1 = [[DocID, ContentItemType, PA, OriginalContentItemPA, DocTitle]]            
        df1 = df1.append(list1) #append list to dataframe, export to csv outside of the loop
        print(list1)
        i=i+1 #increment the counter   
    df1.columns = ["DocID", "ContentItemType", "PA", "OriginalContentItemPA", "DocTitle"]   
    df1.to_csv(ReportType + '.csv', sep=',',index=False) 
    return df1
            

dfUpdateHighlights = DFCleanup(pd.read_csv(weeklyUpdateReportFilename), ShortcutTypeList, reportDir + 'update')
dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilename), ShortcutTypeList, reportDir + 'new')

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')
highlightDate = str(time.strftime("%d %B %Y"))
highlightType = 'weekly'

for PA in AllPAs:
    constantPA = PA        
     
    khdoc = ET.Element('kh:document')
    khbody = ET.SubElement(khdoc, 'kh:body')
    khdoctitle = ET.SubElement(khbody, 'kh:document-title')
    khdoctitle.text = PA + '[weekly/monthly] highlights–[dd Month yyyy]'
    khminisummary = ET.SubElement(khbody, 'kh:mini-summary')
    khminisummary.text = "This [week's/month's] edition of [PA] [weekly/monthly] highlights includes:"
    #News
    trsecmain = ET.SubElement(khbody, 'tr:secmain')
    coretitle = ET.SubElement(trsecmain, 'core:title')
    coretitle.text = '[Subtopic provided]'
    trsecsub1 = ET.SubElement(trsecmain, 'tr:secsub1')
    coretitle = ET.SubElement(trsecsub1, 'core:title')
    coretitle.text = '[News analysis name]'
    corepara = ET.SubElement(trsecsub1, 'core:para')
    corepara.text = '[Mini-summary]'
    corepara = ET.SubElement(trsecsub1, 'core:para')
    corepara.text = 'See News Analysis: [XML ref for News Analysis].'
    #Updated and New content
    trsecmain = ET.SubElement(khbody, 'tr:secmain')
    coretitle = ET.SubElement(trsecmain, 'core:title')
    coretitle.text = 'New and updated content'
    
    #New
    if len(dfNewHighlights[(dfNewHighlights.PA ==PA)]) > 0: #if there are any new docs for the PA, create new doc section
        trsecsub1 = ET.SubElement(trsecmain, 'tr:secsub1')
        coretitle = ET.SubElement(trsecsub1, 'core:title')
        coretitle.text = 'New content'
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']
        for ContentType in ContentTypeList:     
            dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == ContentType)] 
            #dfNew = dfNewHighlights[dfNewHighlights.PA.isin([PA])] 
            newHighlightCount = len(dfNew)
            print(PA, ContentType, newHighlightCount, 'new docs')   
            if ContentType == 'Precedent': ContentTypeHeader = 'Precedents'
            if ContentType == 'PracticeNote': ContentTypeHeader = 'Practice Notes'
            if ContentType == 'Checklist': ContentTypeHeader = 'Checklists'

            if newHighlightCount > 0:    
                generichd2 = ET.SubElement(trsecsub1, 'core:generic-hd-2')
                generichd2.text = ContentTypeHeader   
                for x in range(0,newHighlightCount):   
                    
                    DocID = dfNew.DocID.iloc[x]     
                    DocTitle = dfNew.DocTitle.iloc[x]
                    ContentType = dfNew.ContentItemType.iloc[x]
                    #print(DocID, DocTitle, ContentType)
                    
                    #if ContentTypeList.count(ContentType) < 1 : #Check if ContentType not in the list of valid types                        
                    #    ContentType = dfNew.OriginalContentItemType.iloc[x]
                    #    DocID = str(round(dfNew.OriginalContentItemId.iloc[x],0))
                    OriginalPA = dfNew.OriginalContentItemPA.iloc[x]
                    
                    if ContentType == 'Checklist': ContentType = 'PracticeNote'

                    dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
                    pguidlookup = pguidlistDir + dpsi + '.xml'
                    pguidlookupdom = ET.parse(pguidlistDir + dpsi + '.xml')
                    #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
                    tag = pguidlookupdom.find(".//document[@database-id='" + str(DocID) + "']")
                    pguid = tag.get('pguid')
                    
                    #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)
                    
                    dpsidocID = str(DocID)
                    doctitle = str(DocTitle)
                    
                    corepara = ET.SubElement(trsecsub1, 'core:para')
                    lncicite = ET.SubElement(corepara, 'lnci:cite')
                    lncicite.set('normcite', pguid)
                    lncicontent = ET.SubElement(lncicite, 'lnci:content')
                    lncicontent.text = doctitle

    #Updated  
    if len(dfUpdateHighlights[(dfUpdateHighlights.PA ==PA)]) > 0: #if there are any updated docs for the PA, create updated doc section  
        trsecsub1 = ET.SubElement(trsecmain, 'tr:secsub1')
        coretitle = ET.SubElement(trsecsub1, 'core:title')
        coretitle.text = 'Updated content'
        ContentTypeList=['Precedent', 'PracticeNote', 'Checklist']#, 'QandAs']
        for ContentType in ContentTypeList:    
            dfUpdate = dfUpdateHighlights[(dfUpdateHighlights.PA ==PA) & (dfUpdateHighlights.ContentItemType == ContentType)] 
            updateHighlightCount = len(dfUpdate)
            print(PA, ContentType, updateHighlightCount, 'updates') 
            if ContentType == 'Precedent': ContentTypeHeader = 'Precedents'
            if ContentType == 'PracticeNote': ContentTypeHeader = 'Practice Notes'
            if ContentType == 'Checklist': ContentTypeHeader = 'Checklists'

            if updateHighlightCount > 0:    
                generichd2 = ET.SubElement(trsecsub1, 'core:generic-hd-2')
                generichd2.text = ContentTypeHeader   

                for x in range(0,updateHighlightCount):   
                    DocID = dfUpdate.DocID.iloc[x]     
                    DocTitle = dfUpdate.DocTitle.iloc[x]
                    ContentType = dfUpdate.ContentItemType.iloc[x]
                    #print(DocID, DocTitle, ContentType)
                    
                    #if ContentTypeList.count(ContentType) < 1 : #Check if ContentType not in the list of valid types                        
                    #    ContentType = dfUpdate.OriginalContentItemType.iloc[x]
                    #    DocID = str(round(dfUpdate.OriginalContentItemId.iloc[x],0))
                    OriginalPA = dfUpdate.OriginalContentItemPA.iloc[x]
                    
                    if ContentType == 'Checklist': ContentType = 'PracticeNote'

                    dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
                    pguidlookup = pguidlistDir + dpsi + '.xml'
                    pguidlookupdom = ET.parse(pguidlistDir + dpsi + '.xml')
                    #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
                    #print(DocID, PA, ContentType, dpsi, pguidlookup)
                    tag = pguidlookupdom.find(".//document[@database-id='" + str(DocID) + "']")
                    pguid = tag.get('pguid')
                    
                    #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)
                    
                    dpsidocID = str(DocID)
                    doctitle = str(DocTitle)
                    
                    corepara = ET.SubElement(trsecsub1, 'core:para')
                    lncicite = ET.SubElement(corepara, 'lnci:cite')
                    lncicite.set('normcite', pguid)
                    lncicontent = ET.SubElement(lncicite, 'lnci:content')
                    lncicontent.text = doctitle
    
    #Dates for your diary
    trsecmain = ET.SubElement(khbody, 'tr:secmain')
    coretitle = ET.SubElement(trsecmain, 'core:title')
    coretitle.text = 'Dates for your diary'
    corepara = ET.SubElement(trsecmain, 'core:para')
    corepara.text = 'Table of webinars to go here'
    corepara = ET.SubElement(trsecmain, 'core:para')
    corepara.text = 'A separate subscription is required to view webinars. To book, see: <core:url address="www.lexiswebinars.co.uk/">Webinars</core:url> or please email: <core:url address="webinars@lexisnexis.co.uk" type="mailto">webinars@lexisnexis.co.uk</core:url> quoting the title of the webinar.'
    #Trackers
    trsecmain = ET.SubElement(khbody, 'tr:secmain')
    coretitle = ET.SubElement(trsecmain, 'core:title')
    coretitle.text = 'Trackers'
    corepara = ET.SubElement(trsecmain, 'core:para')
    corepara.text = 'We have the following trackers:'
    corelist = ET.SubElement(trsecmain, 'core:list')
    corelist.set('type', 'bullet')
    corelistitem = ET.SubElement(corelist, 'core:listitem')
    corepara = ET.SubElement(corelistitem, 'core:para')
    lncicite = ET.SubElement(corepara, 'lnci:cite')
    lncicite.set('normcite', dpsidocID)
    lncicontent = ET.SubElement(lncicite, 'lnci:content')
    lncicontent.text = doctitle
    #Latest QAs
   
    #QAs should only ever be new
        
    dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == 'QandAs')] 
    newHighlightCount = len(dfNew)
    print(PA, newHighlightCount, 'new QAs')     
    if newHighlightCount > 0:            
        trsecmain = ET.SubElement(khbody, 'tr:secmain')
        coretitle = ET.SubElement(trsecmain, 'core:title')
        if newHighlightCount > 1:
            coretitle.text = 'Latest Q&amp;As'
        else:
            coretitle.text = 'Latest Q&amp;A'

        for x in range(0,newHighlightCount):               
            DocID = dfNew.DocID.iloc[x]     
            DocTitle = dfNew.DocTitle.iloc[x]
            ContentType = dfNew.ContentItemType.iloc[x]
            OriginalPA = dfNew.OriginalContentItemPA.iloc[x]
            dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
            pguidlookup = pguidlistDir + dpsi + '.xml'
            pguidlookupdom = ET.parse(pguidlistDir + dpsi + '.xml')
            #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
            tag = pguidlookupdom.find(".//document[@database-id='" + str(DocID) + "']")
            pguid = tag.get('pguid')                
            #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)                
            dpsidocID = str(DocID)
            doctitle = str(DocTitle)
            
            corepara = ET.SubElement(trsecmain, 'core:para')
            lncicite = ET.SubElement(corepara, 'lnci:cite')
            lncicite.set('normcite', pguid)
            lncicontent = ET.SubElement(lncicite, 'lnci:content')
            lncicontent.text = doctitle




    #Useful information
    trsecmain = ET.SubElement(khbody, 'tr:secmain')
    coretitle = ET.SubElement(trsecmain, 'core:title')
    coretitle.text = 'Useful information'
    corepara = ET.SubElement(trsecmain, 'core:para')
    corepara.text = 'Automation of these highlights documents is ongoing...'



    tree = ET.ElementTree(khdoc)
    xmlfilepath = 'xml\\' + constantPA + '_highlights.xml'
    tree.write(xmlfilepath,encoding='utf-8')

    f = open(xmlfilepath,'r')
    filedata = f.read()
    f.close()
    newdata = filedata  
    newdata = newdata.replace("<kh:document>","""<?xml version="1.0" encoding="UTF-8"?><!--Arbortext, Inc., 1988-2013, v.4002--><!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd"><?Pub EntList mdash reg #8364 #176 #169 #8230 #10003 #x2610 #x2611 #x2612 #x2613?><?Pub Inc?><kh:document xmlns:core="http://www.lexisnexis.com/namespace/sslrp/core" xmlns:fn="http://www.lexisnexis.com/namespace/sslrp/fn" xmlns:header="http://www.lexisnexis.com/namespace/uk/header" xmlns:kh="http://www.lexisnexis.com/namespace/uk/kh" xmlns:lnb="http://www.lexisnexis.com/namespace/uk/lnb" xmlns:lnci="http://www.lexisnexis.com/namespace/common/lnci" xmlns:tr="http://www.lexisnexis.com/namespace/sslrp/tr">""")
    newdata = newdata.replace("[PA]", constantPA)
    newdata = newdata.replace("[weekly/monthly]", " " + highlightType)
    newdata = newdata.replace("[dd Month yyyy]", highlightDate)
    if highlightType == 'weekly':
        newdata = newdata.replace("[week's/month's]", "week's")
    if highlightType == 'monthly':
        newdata = newdata.replace("[week's/month's]", "month's")
    f = open(xmlfilepath,'w')
    f.write(newdata)
    f.close()

    print('XML exported to...' + xmlfilepath)
    


print('Finished')
