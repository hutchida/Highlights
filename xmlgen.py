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
import shutil

def DFCleanup(df, ShortcutTypeList, ReportType):
    log('Entering DFCleanup...')
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
    log('DFCleanup complete...')
    return df1

    
def xml_generation(PA, highlightDate, highlightType, dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir):
    constantPA = PA        
    log('\n' + PA)
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
            log(ContentType + ', ' + str(newHighlightCount) + ' new docs')
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
                    corepara.text = '● '
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
            log(ContentType + ', ' + str(updateHighlightCount) + ' updates')
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
                    corepara.text = '● '
                    lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                    lncicite.set('normcite', pguid)
                    lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                    lncicontent.text = doctitle
                    if UnderReview == True:
                        comment = etree.Comment('Doc ID: ' + str(DocID) + ' UNDER REVIEW: ' + str(UnderReview))
                        corepara.append(comment) #add comment after link
    
    newAndUpdateContentSection = trsecmain
   
    #QAs should only ever be new
        
    dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == 'QandAs')] 
    dfNew = dfNew.sort_values(['DocTitle'], ascending = True)
    newHighlightCount = len(dfNew)    
    log(str(newHighlightCount) + ' new QAs')
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
                    log('Not found on pguid look up list: ' + str(DocID) + str(DocTitle))        
                #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)              
            except: pguid = 'notfound'

            dpsidocID = str(DocID)
            doctitle = str(DocTitle)
            
            corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
            corepara.text = '● '
            lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
            lncicite.set('normcite', pguid)
            lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
            lncicontent.text = doctitle
            if UnderReview == True:
                comment = etree.Comment('Doc ID: ' + str(DocID) + ' UNDER REVIEW: ' + str(UnderReview))
                corepara.append(comment) #add comment after link
    QAsSection = trsecmain
    #stick a standalone copy in the archive directory regardless
    tree = etree.ElementTree(khdoc)  
    #create standalone xml as backup
    xmlfilepath = outputDir + constantPA + '\\archive\\'    + constantPA + ' New and Updated content ' + highlightDate + '.xml'             
    tree.write(xmlfilepath, encoding="utf-8", xml_declaration=True, doctype=doctype)

    f = open(xmlfilepath,'r', encoding='utf-8')
    filedata = f.read()
    f.close()
    newdata = filedata  
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

    #insert new and updated content into existing template in the auto gen folder
    templateFilepath = FindMostRecentFile(outputDir + constantPA + '\\Auto-generate XML\\', '*template.xml')
    if templateFilepath == 'na': templateFilepath = FindMostRecentFile(outputDir + constantPA + '\\Auto-generate XML\\', '*draft.xml')
    if templateFilepath != 'na':
        templateFilename = re.search('.*\\\\([^\.]*\..*)',templateFilepath).group(1) 
        templateOutputFilepath = outputDir + constantPA + '\\' + templateFilename #output to one folder up
    
        listOfFiles = ListOfFilesInDirectory(outputDir + constantPA + '\\Auto-generate XML\\', '*.xml')
        
        templateInsertionSuccessful = False #default
        print(templateFilepath)
        templateroot = etree.parse(templateFilepath).getroot() 
        templateKHbody = templateroot.find('.//kh:body', NSMAP)    #find element of elements where insertion will be   
        for trsecmain in templateKHbody.findall('.//tr:secmain', NSMAP):
            coretitle = trsecmain.find('core:title', NSMAP)
            parent = trsecmain.getparent() #set parent 
            coretitle_text = coretitle.text
            print(coretitle_text.lower())
            if coretitle_text.lower() == 'new and updated content':                           
                #try:
                print(parent.index(trsecmain))
                templateKHbody.insert(parent.index(trsecmain), newAndUpdateContentSection) #insert new and updated content at position of existing placeholder, using index
                templateKHbody.remove(trsecmain) #delete existing placeholder
                #except:
                #    print('Problem inserting new and updated sections into the template...')
                #    log('Problem inserting new and updated sections into the template...')
            if coretitle.text == 'New Q&As':         
                #try:       
                print(parent.index(trsecmain))
                templateKHbody.insert(parent.index(trsecmain), QAsSection) #insert QA content at position of existing placeholder, using index
                parent = trsecmain.getparent() #delete existing placeholder
                templateKHbody.remove(trsecmain)
                #except:
                #    print('Problem inserting QA sections into the template...')
                #    log('Problem inserting QA sections into the template...')
        templatetree = etree.ElementTree(templateroot)
        templatetree.write(templateOutputFilepath, encoding="utf-8", xml_declaration=True, doctype=doctype)    
        log('Insertion into template successful: \n' + templateOutputFilepath)
        templateInsertionSuccessful = True

        backup_template_filepath = outputDir + constantPA + '\\archive\\backup_' + templateFilename
        shutil.copy(templateFilepath, backup_template_filepath)
        log('Created backup of dropped template: \n' + backup_template_filepath)

        try:
            for existingFilepath in listOfFiles:
                os.remove(existingFilepath) #Delete old file
                log('Deleted from auto-gen folder: \n' + existingFilepath)
        except:
            log('Problem deleting: \n' + existingFilepath)

    else: 
        log('No template xml file found in auto-gen folder...')
        templateInsertionSuccessful = False 


    if templateInsertionSuccessful == False: 
        #REMOVE ONCE ALL DOING WEEKLY
        if PA in MonthlyPAs:
            xmlfilepath = outputDir + constantPA + '\\weekly\\' + constantPA + ' New and Updated content ' + highlightDate + '.xml'
        else:
            xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' New and Updated content ' + highlightDate + '.xml'

        try: tree.write(xmlfilepath, encoding="utf-8", xml_declaration=True, doctype=doctype)
        except PermissionError: log('Permission error encountered, cannot write to ' + xmlfilepath)

    f = open(xmlfilepath,'r', encoding='utf-8')
    filedata = f.read()
    f.close()
    newdata = filedata  
    #newdata = newdata.replace("<kh:document ","""<?xml version="1.0" encoding="UTF-8"?><!--Arbortext, Inc., 1988-2013, v.4002--><!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd"><?Pub EntList mdash reg #8364 #176 #169 #8230 #10003 #x2610 #x2611 #x2612 #x2613?><?Pub Inc?><kh:document """)#<kh:document xmlns:core="http://www.lexisnexis.com/namespace/sslrp/core" xmlns:fn="http://www.lexisnexis.com/namespace/sslrp/fn" xmlns:header="http://www.lexisnexis.com/namespace/uk/header" xmlns:kh="http://www.lexisnexis.com/namespace/uk/kh" xmlns:lnb="http://www.lexisnexis.com/namespace/uk/lnb" xmlns:lnci="http://www.lexisnexis.com/namespace/common/lnci" xmlns:tr="http://www.lexisnexis.com/namespace/sslrp/tr">""")
    newdata = newdata.replace("[PA]", constantPA)
    newdata = newdata.replace("Life Sciences and Pharmaceuticals", "Life Sciences")
    newdata = newdata.replace("[weekly/monthly]", highlightType)
    newdata = newdata.replace("[dd Month yyyy]", highlightDate)
    if highlightType == 'weekly':
        newdata = newdata.replace("[week's/month's]", "week's")
    if highlightType == 'monthly':
        newdata = newdata.replace("[week's/month's]", "month's")
    try: 
        f = open(xmlfilepath,'w', encoding='utf-8')
        f.write(newdata)
        f.close()
    except PermissionError: log('Permission error encountered, cannot write to ' + xmlfilepath)


    log('Standalone XML exported to: \n' + xmlfilepath + '\n')

def ListOfFilesInDirectory(directory, pattern):
    existingList = []
    filelist = glob.iglob(os.path.join(directory, pattern)) #builds list of file in a directory based on a pattern
    for filepath in filelist: existingList.append(filepath)        
    return existingList

def Archive(listOfFiles, reportDir):
    copyList = []
    archiveDir = reportDir + 'Archive\\'

    #Check archive folder exists
    if os.path.isdir(archiveDir) == False:
        os.makedirs(archiveDir)

    for existingFilepath in listOfFiles:
        print(existingFilepath)
        directory = re.search('(.*\\\\)[^\.]*\.xml',existingFilepath).group(1)
        filename = re.search('.*\\\\([^\.]*\.xml)',existingFilepath).group(1)
        destinationFilepath = directory + 'Archive\\' + filename
        shutil.copy(existingFilepath, destinationFilepath) #Copy
        os.remove(existingFilepath) #Delete old file
        copyList.append('Moved: ' + existingFilepath + ', to: ' + destinationFilepath)
    return copyList


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


def log(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()
    print(message)


pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'

#Updated 090420 1351 - monthly pas now outputting as weeklies but to respective weekly folder
#Updated 170420 1024 - fixed issue with empty standalone xml going to archive folder
# change xmlgen to look for variations of 'new and updated content' header - 230420 1055
# change xmlgen to look for dropped template filename rather than predicted template filename - 230420 1055
# change xmlgen to stop search and replace of xml header, which was corrupting standalone new and updated xml - 230420 1055
# add to xmlgen logging of insertion - 230420 1055
# create backup of template in archive folder - 230420 1055
#290420 1636 - added line to accept *draft.xml in the autogen folder

#state = 'livedev'
state = 'live'

if state == 'live':
    logDir = "\\\\atlas\\lexispsl\\Highlights\\logs\\" 
    reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
    givendate = datetime.datetime.today()
    givenstrdate =  str(time.strftime("%Y-%m-%d"))

if state == 'livedev':
    logDir = "\\\\atlas\\lexispsl\\Highlights\\dev\\logs\\" 
    #reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
    reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'
    givenstrdate =  "2020-4-30"
    givendate = datetime.date(2020, 4, 30)    
    #givendate = datetime.datetime.today()
    givenstrdate =  str(time.strftime("%Y-%m-%d"))

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
#AllPAs = ['Life Sciences and Pharmaceuticals']    
#AllPAs = ['Energy', 'Environment', 'Financial Services', 'Information Law', 'IP', 'Life Sciences and Pharmaceuticals', 'Tax', 'TMT']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
ShortcutTypeList = ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut']
MonthlyDates = ['2019-11-29', '2019-12-20', '2020-01-31', '2020-02-28', '2020-03-31', '2020-04-29', '2020-05-29', '2020-06-30', '2020-07-31', '2020-08-28', '2020-09-30', '2020-10-30', '2020-11-30', '2020-12-18']
doctype = '<!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd">'

#print(givendate)
#print('Is last working day of the month:' + str(IsLastWorkingDayOfMonth(givendate)))
#print('Is a Thursday:' + str(IsThursday(givendate)))    

JCSLogFile = logDir + 'JCSlog-xmlgen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d/%m/%Y %H:%M:%S"))
l.write("Start "+logdate+"\n")
l.close()


#main script
log("Today's date is: " + str(givendate))
if str(givenstrdate) in MonthlyDates:
    log("\nXML auto-generation for monthly highlights...\n")
    monthlyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_monthly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    monthlyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_monthly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    dfUpdateHighlights = DFCleanup(pd.read_csv(monthlyUpdateReportFilepath), ShortcutTypeList, reportDir + 'update')
    dfNewHighlights = DFCleanup(pd.read_csv(monthlyNewReportFilepath), ShortcutTypeList, reportDir + 'new')
    highlightDate = str(time.strftime("%B %Y")) 
    dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')
    for PA in MonthlyPAs:
        xml_generation(PA, highlightDate, 'monthly', dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir)
else: 
    log('Today is not a date in the monthly highlights list of dates...skipping monthly highlight xml generation...')

if IsThursday(givendate) == True:
    log("\nXML auto-generation for weekly highlights...\n")
    weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    weeklyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    dfUpdateHighlights = DFCleanup(pd.read_csv(weeklyUpdateReportFilepath), ShortcutTypeList, reportDir + 'update')
    dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilepath), ShortcutTypeList, reportDir + 'new')
    highlightDate = str(time.strftime("%#d %B %Y")) #the hash character turns off the leading zero in the day
    #highlightDate = '26 March 2020'
    dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')    
    for PA in AllPAs:
        #if PA not in MonthlyPAs:
        xml_generation(PA, highlightDate, 'weekly', dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir)
else: 
    log('Today is not a Thursday...skipping weekly highlight xml generation...')


log('Finished')


#wait = input("PAUSED...when ready press enter")