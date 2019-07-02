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

    #df['LastPublishedDate'] = pd.to_datetime(df['LastPublishedDate'], dayfirst=True) #UK date
    #df['LastUnderReviewDate'] = pd.to_datetime(df['LastUnderReviewDate'], dayfirst=True) #UK date

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

    #df['MajorUpdateFirstPublished'] = pd.to_datetime(df['MajorUpdateFirstPublished'], dayfirst=True) #UK date
    #df['DateFirstPublished'] = pd.to_datetime(df['DateFirstPublished'], dayfirst=True) #UK date

    #df1 = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    
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
                    #dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
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
                            #dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
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


def XMLGenerationWeekly(PA, highlightDate, highlightFileDate, highlightType, dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir):
    constantPA = PA        
    if PA == 'Life Sciences': PA = 'Life Sciences and Pharmaceuticals'
    if PA == 'Share Incentives': PA = 'Share Schemes'
    if PA == 'Insurance and Reinsurance': PA = 'Insurance'
    khdoc = ET.Element('kh:document')
    khbody = ET.SubElement(khdoc, 'kh:body')
    khdoctitle = ET.SubElement(khbody, 'kh:document-title')
    #khdoctitle.text = PA + '[weekly/monthly] highlights—[dd Month yyyy]'
    khdoctitle.text = PA + ' [weekly/monthly] highlights—[dd Month yyyy]'
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
    #if len(dfNewHighlights[(dfNewHighlights.PA ==PA)]) > 0: #if there are any new docs for the PA, create new doc section
    if len(dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType != 'QandAs')]) > 0: #if there are any new docs for the PA, create new doc section
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
    #if len(dfUpdateHighlights[(dfUpdateHighlights.PA ==PA)]) > 0: #if there are any updated docs for the PA, create updated doc section 
    if len(dfUpdateHighlights[(dfUpdateHighlights.PA ==PA) & (dfUpdateHighlights.ContentItemType != 'QandAs')]) > 0: #if there are any new docs for the PA, create new doc section 
        trsecsub1 = ET.SubElement(trsecmain, 'tr:secsub1')
        coretitle = ET.SubElement(trsecsub1, 'core:title')
        coretitle.text = 'Updated content'
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']#, 'QandAs']
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
    #corepara = ET.SubElement(trsecmain, 'core:para')
    #corepara.text = 'A separate subscription is required to view webinars. To book, see: <core:url address="www.lexiswebinars.co.uk/">Webinars</core:url> or please email: <core:url address="webinars@lexisnexis.co.uk" type="mailto">webinars@lexisnexis.co.uk</core:url> quoting the title of the webinar.'
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
    lncicite.set('normcite', 'pguid-link-goes-here')
    lncicontent = ET.SubElement(lncicite, 'lnci:content')
    lncicontent.text = 'tracker doc title goes here'
    #Latest QAs
   
    #QAs should only ever be new
        
    dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == 'QandAs')] 
    newHighlightCount = len(dfNew)
    print(PA, newHighlightCount, 'new QAs')     
    if newHighlightCount > 0:            
        trsecmain = ET.SubElement(khbody, 'tr:secmain')
        coretitle = ET.SubElement(trsecmain, 'core:title')
        if newHighlightCount > 1:
            coretitle.text = 'Latest Q&As'
        else:
            coretitle.text = 'Latest Q&A'

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
    xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' Weekly highlights ' + highlightFileDate + ' test.xml'
    tree.write(xmlfilepath,encoding='utf-8')

    f = open(xmlfilepath,'r')
    filedata = f.read()
    f.close()
    newdata = filedata  
    newdata = newdata.replace("<kh:document>","""<?xml version="1.0" encoding="UTF-8"?><!--Arbortext, Inc., 1988-2013, v.4002--><!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd"><?Pub EntList mdash reg #8364 #176 #169 #8230 #10003 #x2610 #x2611 #x2612 #x2613?><?Pub Inc?><kh:document xmlns:core="http://www.lexisnexis.com/namespace/sslrp/core" xmlns:fn="http://www.lexisnexis.com/namespace/sslrp/fn" xmlns:header="http://www.lexisnexis.com/namespace/uk/header" xmlns:kh="http://www.lexisnexis.com/namespace/uk/kh" xmlns:lnb="http://www.lexisnexis.com/namespace/uk/lnb" xmlns:lnci="http://www.lexisnexis.com/namespace/common/lnci" xmlns:tr="http://www.lexisnexis.com/namespace/sslrp/tr">""")
    newdata = newdata.replace("[PA]", constantPA)
    newdata = newdata.replace("[weekly/monthly]", highlightType)
    newdata = newdata.replace("[dd Month yyyy]", highlightDate)
    if highlightType == 'weekly':
        newdata = newdata.replace("[week's/month's]", "week's")
    if highlightType == 'monthly':
        newdata = newdata.replace("[week's/month's]", "month's")
    f = open(xmlfilepath,'w')
    f.write(newdata)
    f.close()

    print('XML exported to...' + xmlfilepath)


def FindMostRecentFile(directory, pattern):
    filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
    filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
    return filelist[0]



#main script
print("New and Updated content report for highlights...\n\n")

#Directories
reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
#reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
#aicerDir = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"
aicerDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER_AM\\'
#aicerDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER\\'
#aicerDir = '\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER_PM\\'
globalmetricsDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Global_Metrics\\'
#globalmetricsDir = '\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER_withShortcut_AdHoc\\'
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'
#outputDir = 'xml\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'


aicerFilename = FindMostRecentFile(aicerDir, '*AICER*.csv')
#aicerFilename = FindMostRecentFile(aicerDir, '*AllContentItemsExport_[0-9][0-9][0-9][0-9].csv')
#aicerFilename = re.search('.*\\\\AICER\\\\([^\.]*\.csv)',aicerFilename).group(1)
aicerFilename = re.search('.*\\\\AICER_AM\\\\([^\.]*\.csv)',aicerFilename).group(1)
#aicerFilename = re.search('.*\\\\AICER_PM\\\\([^\.]*\.csv)',aicerFilename).group(1)
print('Loading the most recent AICER report: ' + aicerFilename)
#filter
dfaicer = pd.read_csv(aicerDir + aicerFilename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
print('Aicer loaded...loading Aicer shortcuts...')
dfshortcuts =  pd.read_csv(globalmetricsDir + 'AllContentItemsExportWithShortCutNodeInfo.csv', encoding='utf-8', low_memory=False) #Load csv file into dataframe
print('Aicer shortcuts loaded...filtering reports...')

#Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'weekly', 'new')
#Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'weekly', 'updated')
#Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'monthly', 'new')
#Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'monthly', 'updated')

wait = input("PAUSED...when ready press enter")

#weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AllContentItemsExport_[0-9][0-9]_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
#weeklyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AllContentItemsExport_[0-9][0-9]_UKPSL_weekly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
#monthlyNewReportFilepath = FindMostRecentFile(reportDir, '*AllContentItemsExport_[0-9][0-9]_UKPSL_monthly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
#monthlyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AllContentItemsExport_[0-9][0-9]_UKPSL_monthly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')


weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
weeklyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
monthlyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_monthly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
monthlyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_monthly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')


#Cleanup the highlights list ready for xml generation
ShortcutTypeList = ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut']
dfUpdateHighlights = DFCleanup(pd.read_csv(weeklyUpdateReportFilepath), ShortcutTypeList, reportDir + 'update')
dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilepath), ShortcutTypeList, reportDir + 'new')

#XML generation for all PAs
AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance and Reinsurance', 'IP', 'Life Sciences', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Incentives', 'Tax', 'TMT', 'Wills and Probate']    
dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')
highlightDate = str(time.strftime("%#d %B %Y")) #the hash character turns off the leading zero in the day
#highlightFileDate = str(time.strftime("%d%m%Y"))
highlightFileDate = str(time.strftime("%#d %B %Y"))
highlightType = 'weekly'

for PA in AllPAs:
    XMLGenerationWeekly(PA, highlightDate, highlightFileDate, highlightType, dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir)

print('Finished')
