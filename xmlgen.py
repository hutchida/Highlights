#XML generator for highlights, loops through all entries on the new and updated content lists and creates xml with links to those mentioned documents that can be copied and pasted into the highlights templates
#Developed by Daniel Hutchings

import csv
import pandas as pd
from pandas.tseries.offsets import BMonthEnd
import glob
import os
import fnmatch
import re
import time
import datetime
import sys
import xml.etree.ElementTree as ET
from lxml import etree

def DFCleanup(df, ShortcutTypeList, ReportType):
    LogOutput('Entering DFCleanup...')
    df1 = pd.DataFrame()
    i=0
    for index, row in df.iterrows(): #loop through df deconstruct/reconstruct
        DocID = df.DocID.iloc[i]
        ContentItemType = df.ContentItemType.iloc[i]
        PA = df.PA.iloc[i]
        OriginalContentItemPA = PA
        DocTitle = df.DocTitle.iloc[i]
        UnderReview = df.UnderReview.iloc[i]
        if ShortcutTypeList.count(ContentItemType) > 0: #if found in list
            ContentItemType = df.OriginalContentItemType.iloc[i]
            DocID = df.OriginalContentItemId.iloc[i]
            OriginalContentItemPA = df.OriginalContentItemPA.iloc[i]

        list1 = [[DocID, ContentItemType, PA, OriginalContentItemPA, DocTitle, UnderReview]]            
        df1 = df1.append(list1) #append list to dataframe, export to csv outside of the loop
        #print(list1)
        i=i+1 #increment the counter   
    df1.columns = ["DocID", "ContentItemType", "PA", "OriginalContentItemPA", "DocTitle", "UnderReview"]   
    #df1.to_csv(ReportType + '.csv', sep=',',index=False) 
    LogOutput('DFCleanup complete...')
    return df1

    
def XMLGenerationWeekly(PA, highlightDate, highlightType, dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir):
    constantPA = PA        
    
    NSMAP = {'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}

    if PA == 'Life Sciences': PA = 'Life Sciences and Pharmaceuticals'
    if PA == 'Share Incentives': PA = 'Share Schemes'
    if PA == 'Insurance and Reinsurance': PA = 'Insurance'
    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
    khbody = etree.SubElement(khdoc, '{%s}body' % NSMAP['kh'])
    
    khdoctitle = etree.SubElement(khbody, '{%s}document-title' % NSMAP['kh'])
    khdoctitle.text = PA + ' [weekly/monthly] highlights—[dd Month yyyy]'
    khminisummary = etree.SubElement(khbody, '{%s}mini-summary' % NSMAP['kh'])
    khminisummary.text = "This [week's/month's] edition of [PA] [weekly/monthly] highlights includes:"

    #News
    #trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    #coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    #coretitle.text = '[Subtopic provided]'
    #trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
    #coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
    #coretitle.text = '[News analysis name]'
    #corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
    #corepara.text = '[Mini-summary]'
    #corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
    #corepara.text = 'See News Analysis: [XML ref for News Analysis].'
    
    #Updated and New content
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'New and updated content'
    
    #New
    if len(dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType != 'QandAs')]) > 0: #if there are any new docs for the PA, create new doc section
        
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']
        for ContentType in ContentTypeList:     
            dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == ContentType)] 
            dfNew = dfNew.sort_values(['DocTitle'], ascending = True)
            newHighlightCount = len(dfNew)
            print(PA, ContentType, newHighlightCount, 'new docs')   
            LogOutput(PA + ', ' + ContentType + ', ' + str(newHighlightCount) + ' new docs')
            if ContentType == 'Precedent':
                if newHighlightCount > 1: ContentTypeHeader = 'New Precedents'
                else: ContentTypeHeader = 'New Precedent'

            if ContentType == 'PracticeNote':
                if newHighlightCount > 1: ContentTypeHeader = 'New Practice Notes'
                else: ContentTypeHeader = 'New Practice Note'

            if ContentType == 'Checklist':
                if newHighlightCount > 1: ContentTypeHeader = 'New Checklists'
                else: ContentTypeHeader = 'New Checklist'

            if newHighlightCount > 0:    
                trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
                coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
                coretitle.text = ContentTypeHeader   
                for x in range(0,newHighlightCount):   
                    
                    DocID = dfNew.DocID.iloc[x]     
                    DocTitle = dfNew.DocTitle.iloc[x]
                    ContentType = dfNew.ContentItemType.iloc[x]
                    OriginalPA = dfNew.OriginalContentItemPA.iloc[x]
                    UnderReview = dfNew.UnderReview.iloc[x]

                    if ContentType == 'Checklist': ContentType = 'PracticeNote'

                    dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
                    #pguidlookup = pguidlistDir + dpsi + '.xml'
                    try:
                        pguidlookupdom = etree.parse(pguidlistDir + dpsi + '.xml')
                        #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
                        tag = pguidlookupdom.find(".//document[@database-id='" + str(DocID) + "']")
                        pguid = tag.get('pguid')
                    except:
                        pguid = 'notfound'
                    
                    #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)
                    
                    dpsidocID = str(DocID)
                    doctitle = str(DocTitle)
                    
                    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                    lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                    lncicite.set('normcite', pguid)
                    lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                    lncicontent.text = doctitle
                    if UnderReview == True:
                        comment = etree.Comment('Doc ID: ' + str(DocID) + ' UNDER REVIEW: ' + str(UnderReview))
                        corepara.append(comment) #add comment after link

    #Updated  
    if len(dfUpdateHighlights[(dfUpdateHighlights.PA ==PA) & (dfUpdateHighlights.ContentItemType != 'QandAs')]) > 0: #if there are any new docs for the PA, create new doc section 
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']#, 'QandAs']
        for ContentType in ContentTypeList:    
            dfUpdate = dfUpdateHighlights[(dfUpdateHighlights.PA ==PA) & (dfUpdateHighlights.ContentItemType == ContentType)] 
            dfUpdate = dfUpdate.sort_values(['DocTitle'], ascending = True)
            updateHighlightCount = len(dfUpdate)
            print(PA, ContentType, updateHighlightCount, 'updates') 
            LogOutput(PA + ', ' + ContentType + ', ' + str(updateHighlightCount) + ' updates')
            if ContentType == 'Precedent': 
                if updateHighlightCount > 1: ContentTypeHeader = 'Updated Precedents'
                else: ContentTypeHeader = 'Updated Precedent'
            if ContentType == 'PracticeNote':
                if updateHighlightCount > 1: ContentTypeHeader = 'Updated Practice Notes'
                else: ContentTypeHeader = 'Updated Practice Note'
            if ContentType == 'Checklist':
                if updateHighlightCount > 1: ContentTypeHeader = 'Updated Checklists'
                else: ContentTypeHeader = 'Updated Checklist'

            if updateHighlightCount > 0:    
                trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
                coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
                coretitle.text = ContentTypeHeader   

                for x in range(0,updateHighlightCount):   
                    DocID = dfUpdate.DocID.iloc[x]     
                    DocTitle = dfUpdate.DocTitle.iloc[x]
                    ContentType = dfUpdate.ContentItemType.iloc[x]
                    OriginalPA = dfUpdate.OriginalContentItemPA.iloc[x]
                    UnderReview = dfUpdate.UnderReview.iloc[x]
                    
                    if ContentType == 'Checklist': ContentType = 'PracticeNote'

                    dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
                    
                    
                    #pguidlookup = pguidlistDir + dpsi + '.xml'
                    try:
                        pguidlookupdom = etree.parse(pguidlistDir + dpsi + '.xml')
                        #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
                        #print(DocID, PA, ContentType, dpsi, pguidlookup)
                        tag = pguidlookupdom.find(".//document[@database-id='" + str(DocID) + "']")
                        pguid = tag.get('pguid')
                    except:
                        pguid = 'notfound'
                    #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)
                    
                    dpsidocID = str(DocID)
                    doctitle = str(DocTitle)
                    
                    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                    lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                    lncicite.set('normcite', pguid)
                    lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                    lncicontent.text = doctitle
                    if UnderReview == True:
                        comment = etree.Comment('Doc ID: ' + str(DocID) + ' UNDER REVIEW: ' + str(UnderReview))
                        corepara.append(comment) #add comment after link
    
    
   
    #QAs should only ever be new
        
    dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == 'QandAs')] 
    dfNew = dfNew.sort_values(['DocTitle'], ascending = True)
    newHighlightCount = len(dfNew)
    print(PA, newHighlightCount, 'new QAs')     
    LogOutput(PA + ', ' + str(newHighlightCount) + ' new QAs')
    if newHighlightCount > 0:            
        trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
        coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
        if newHighlightCount > 1:
            coretitle.text = 'Latest Q&As'
        else:
            coretitle.text = 'Latest Q&A'

        for x in range(0,newHighlightCount):               
            DocID = dfNew.DocID.iloc[x]     
            DocTitle = dfNew.DocTitle.iloc[x]
            ContentType = dfNew.ContentItemType.iloc[x]
            OriginalPA = dfNew.OriginalContentItemPA.iloc[x]
            UnderReview = dfNew.UnderReview.iloc[x]
            dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
            #pguidlookup = pguidlistDir + dpsi + '.xml'
            try:
                pguidlookupdom = etree.parse(pguidlistDir + dpsi + '.xml')
                #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
                tag = pguidlookupdom.find(".//document[@database-id='" + str(DocID) + "']")
                try: pguid = tag.get('pguid')    
                except AttributeError: 
                    pguid = 'notfound'
                    print('Not found on pguid look up list: ' + str(DocID) + str(DocTitle)) 
                    LogOutput('Not found on pguid look up list: ' + str(DocID) + str(DocTitle))        
                #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)              
            except: pguid = 'notfound'

            dpsidocID = str(DocID)
            doctitle = str(DocTitle)
            
            corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
            lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
            lncicite.set('normcite', pguid)
            lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
            lncicontent.text = doctitle
            if UnderReview == True:
                comment = etree.Comment('Doc ID: ' + str(DocID) + ' UNDER REVIEW: ' + str(UnderReview))
                corepara.append(comment) #add comment after link

    

    tree = etree.ElementTree(khdoc)
    xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' New and Updated content ' + highlightDate + ' test.xml'
    tree.write(xmlfilepath,encoding='utf-8')

    f = open(xmlfilepath,'r', encoding='utf-8')
    filedata = f.read()
    f.close()
    newdata = filedata  
    newdata = newdata.replace("<kh:document ","""<?xml version="1.0" encoding="UTF-8"?><!--Arbortext, Inc., 1988-2013, v.4002--><!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd"><?Pub EntList mdash reg #8364 #176 #169 #8230 #10003 #x2610 #x2611 #x2612 #x2613?><?Pub Inc?><kh:document """)#<kh:document xmlns:core="http://www.lexisnexis.com/namespace/sslrp/core" xmlns:fn="http://www.lexisnexis.com/namespace/sslrp/fn" xmlns:header="http://www.lexisnexis.com/namespace/uk/header" xmlns:kh="http://www.lexisnexis.com/namespace/uk/kh" xmlns:lnb="http://www.lexisnexis.com/namespace/uk/lnb" xmlns:lnci="http://www.lexisnexis.com/namespace/common/lnci" xmlns:tr="http://www.lexisnexis.com/namespace/sslrp/tr">""")
    newdata = newdata.replace("[PA]", constantPA)
    newdata = newdata.replace("Life Sciences and Pharmaceuticals", "Life Sciences")
    newdata = newdata.replace("[weekly/monthly]", highlightType)
    newdata = newdata.replace("[dd Month yyyy]", highlightDate)
    if highlightType == 'weekly':
        newdata = newdata.replace("[week's/month's]", "week's")
    if highlightType == 'monthly':
        newdata = newdata.replace("[week's/month's]", "month's")
    f = open(xmlfilepath,'w', encoding='utf-8')
    f.write(newdata)
    f.close()

    print('XML exported to...' + xmlfilepath)
    LogOutput('XML exported to...' + xmlfilepath)
   

def FindMostRecentFile(directory, pattern):
    try:
        filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
        filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
        return filelist[0]
    except: return 'na'


def IsLastWorkingDayOfMonth(givendate):
    offset = BMonthEnd()
    #givendate += datetime.timedelta(days=3)
    lastday = offset.rollforward(givendate)
    if givendate == lastday.date(): return True
    else: return False 

def IsThursday(givendate):
    print(givendate.weekday())
    if givendate.weekday() == 3: return True
    else: return False


def LogOutput(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()


#Directories
#reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
#reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'
logDir = "\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Logs\\"    
#outputDir = 'xml\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
ShortcutTypeList = ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut']
MonthlyDates = ['2019-11-29', '2020-12-20', '2020-02-28', '2020-03-31', '2020-04-29', '2020-05-29', '2020-06-30', '2020-07-31', '2020-08-28', '2020-09-30', '2020-10-30', '2020-11-30', '2020-12-18']
#givendate is auto generated by datetime.datetime.today(), but if you need to manually override it you can set it with datetime.date(year, month, day)
givendate = datetime.datetime.today()
givenstrdate =  str(time.strftime("%Y-%m-%d"))
#givendate = datetime.date(2019, 11, 29)

#print(givendate)
#print('Is last working day of the month:' + str(IsLastWorkingDayOfMonth(givendate)))
#print('Is a Thursday:' + str(IsThursday(givendate)))    

JCSLogFile = logDir + 'JCSlog-xmlgen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()


#main script
print("Today's date is: " + str(givendate))
LogOutput("Today's date is: " + str(givendate))
if str(givenstrdate) in MonthlyDates:
    print("\nXML auto-generation for monthly highlights...\n")
    LogOutput("\nXML auto-generation for monthly highlights...\n")
    monthlyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_monthly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    monthlyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_monthly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    dfUpdateHighlights = DFCleanup(pd.read_csv(monthlyUpdateReportFilepath), ShortcutTypeList, reportDir + 'update')
    dfNewHighlights = DFCleanup(pd.read_csv(monthlyNewReportFilepath), ShortcutTypeList, reportDir + 'new')
    highlightDate = str(time.strftime("%B %Y")) 
    dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')
    for PA in MonthlyPAs:
        XMLGenerationWeekly(PA, highlightDate, 'monthly', dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir)
else: 
    print('Today is not a date in the monthly highlights list of dates...skipping monthly highlight xml generation...')
    LogOutput('Today is not a date in the monthly highlights list of dates...skipping monthly highlight xml generation...')

if IsThursday(givendate) == True:
    print("\nXML auto-generation for weekly highlights...\n")
    LogOutput("\nXML auto-generation for weekly highlights...\n")
    weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    weeklyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    dfUpdateHighlights = DFCleanup(pd.read_csv(weeklyUpdateReportFilepath), ShortcutTypeList, reportDir + 'update')
    dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilepath), ShortcutTypeList, reportDir + 'new')
    highlightDate = str(time.strftime("%#d %B %Y")) #the hash character turns off the leading zero in the day
    dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')    
    for PA in AllPAs:
        if PA not in MonthlyPAs:
            XMLGenerationWeekly(PA, highlightDate, 'weekly', dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir)
else: 
    print('Today is not a Thursday...skipping weekly highlight xml generation...')
    LogOutput('Today is not a Thursday...skipping weekly highlight xml generation...')



print('Finished')
LogOutput('Finished')


#wait = input("PAUSED...when ready press enter")