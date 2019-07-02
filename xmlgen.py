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
        #print(list1)
        i=i+1 #increment the counter   
    df1.columns = ["DocID", "ContentItemType", "PA", "OriginalContentItemPA", "DocTitle"]   
    #df1.to_csv(ReportType + '.csv', sep=',',index=False) 
    return df1

#get rid of namespaces
def strip_ns_prefix(tree):
    #xpath query for selecting all element nodes in namespace
    query = "descendant-or-self::*[namespace-uri()!='']"
    #for each element returned by the above xpath query...
    for element in tree.xpath(query):
        #replace element name with its local name
        element.tag = etree.QName(element).localname
    return tree
    
def XMLGenerationWeekly(PA, highlightDate, highlightFileDate, highlightType, dfdpsi, dfUpdateHighlights, dfNewHighlights, outputDir):
    constantPA = PA        
    PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*' + constantPA + ' Weekly highlights *.xml')
    print('lasthighlightsfilepath: ' + PrevHighlightsFilepath)
    #extract info from last highlights' doc
    NSMAP = {'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}
    try:
        tree = etree.parse(PrevHighlightsFilepath)
        #strip_ns_prefix(tree)
        root = tree.getroot()


        trsecmains = root.findall('.//tr:secmain', NSMAP)
        for trsecmain in trsecmains:
            coretitle = trsecmain.find('core:title', NSMAP)
            try:
                del coretitle.attrib["id"]
                print('\nDELETED id attribute...' + coretitle.text)
            except KeyError: pass
                
            try: 
                if coretitle.text == 'Dates for your diary': DatesSection = trsecmain
            except: pass
            try: 
                if coretitle.text == 'Trackers': TrackersSection = trsecmain
            except: pass
            try: 
                if coretitle.text == 'Useful information': UsefulInfoSection = trsecmain
            except: pass
    except:
        print('Problem loading previous xml for: ' + constantPA)


    if PA == 'Life Sciences': PA = 'Life Sciences and Pharmaceuticals'
    if PA == 'Share Incentives': PA = 'Share Schemes'
    if PA == 'Insurance and Reinsurance': PA = 'Insurance'
    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)#('document', nsmap=NSMAP)
    khbody = etree.SubElement(khdoc, '{%s}body' % NSMAP['kh'])
    
    khdoctitle = etree.SubElement(khbody, '{%s}document-title' % NSMAP['kh'])
    #khdoctitle.text = PA + '[weekly/monthly] highlights—[dd Month yyyy]'
    khdoctitle.text = PA + ' [weekly/monthly] highlights—[dd Month yyyy]'
    khminisummary = etree.SubElement(khbody, '{%s}mini-summary' % NSMAP['kh'])
    khminisummary.text = "This [week's/month's] edition of [PA] [weekly/monthly] highlights includes:"

    #News
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = '[Subtopic provided]'
    trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
    coretitle.text = '[News analysis name]'
    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
    corepara.text = '[Mini-summary]'
    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
    corepara.text = 'See News Analysis: [XML ref for News Analysis].'
    #Updated and New content
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'New and updated content'
    
    #New
    if len(dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType != 'QandAs')]) > 0: #if there are any new docs for the PA, create new doc section
        
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']
        for ContentType in ContentTypeList:     
            dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == ContentType)] 
            newHighlightCount = len(dfNew)
            print(PA, ContentType, newHighlightCount, 'new docs')   
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

    #Updated  
    if len(dfUpdateHighlights[(dfUpdateHighlights.PA ==PA) & (dfUpdateHighlights.ContentItemType != 'QandAs')]) > 0: #if there are any new docs for the PA, create new doc section 
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']#, 'QandAs']
        for ContentType in ContentTypeList:    
            dfUpdate = dfUpdateHighlights[(dfUpdateHighlights.PA ==PA) & (dfUpdateHighlights.ContentItemType == ContentType)] 
            updateHighlightCount = len(dfUpdate)
            print(PA, ContentType, updateHighlightCount, 'updates') 
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
    
    #Dates for your diary
    
    try: khbody.append(DatesSection)
    except: print('No Dates for your diary section found...')
    
    #trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    #coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    #coretitle.text = 'Dates for your diary'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'Table of webinars to go here'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'A separate subscription is required to view webinars. To book, see: <core:url address="www.lexiswebinars.co.uk/">Webinars</core:url> or please email: <core:url address="webinars@lexisnexis.co.uk" type="mailto">webinars@lexisnexis.co.uk</core:url> quoting the title of the webinar.'
    
    #Trackers

    
    try: khbody.append(TrackersSection)
    except: print('No Tracker section found...')
    
    #trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    #coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    #coretitle.text = 'Trackers'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'We have the following trackers:'
    #corelist = etree.SubElement(trsecmain, '{%s}list' % NSMAP['core'])
    #corelist.set('type', 'bullet')
    #corelistitem = etree.SubElement(corelist, '{%s}listitem' % NSMAP['core'])
    #corepara = etree.SubElement(corelistitem, '{%s}para' % NSMAP['core'])
    #lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
    #lncicite.set('normcite', 'pguid-link-goes-here')
    #lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
    #lncicontent.text = 'tracker doc title goes here'
    #Latest QAs
   
    #QAs should only ever be new
        
    dfNew = dfNewHighlights[(dfNewHighlights.PA ==PA) & (dfNewHighlights.ContentItemType == 'QandAs')] 
    newHighlightCount = len(dfNew)
    print(PA, newHighlightCount, 'new QAs')     
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
                #print(DocID, DocTitle, ContentType, dpsi, pguidlookup, pguid)              
            except: pguid = 'notfound'

            dpsidocID = str(DocID)
            doctitle = str(DocTitle)
            
            corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
            lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
            lncicite.set('normcite', pguid)
            lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
            lncicontent.text = doctitle

    #Useful information
    try: khbody.append(UsefulInfoSection)
    except: print('No Useful information section found...')
    #trsecmain = etree.SubElement(khbody, 'tr:secmain')
    #coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    #coretitle.text = 'Useful information'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'Automation of these highlights documents is ongoing...'



    tree = etree.ElementTree(khdoc)
    xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' Weekly highlights ' + highlightFileDate + ' test.xml'
    tree.write(xmlfilepath,encoding='utf-8')

    f = open(xmlfilepath,'r', encoding='utf-8')
    filedata = f.read()
    f.close()
    newdata = filedata  
    newdata = newdata.replace("<kh:document ","""<?xml version="1.0" encoding="UTF-8"?><!--Arbortext, Inc., 1988-2013, v.4002--><!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd"><?Pub EntList mdash reg #8364 #176 #169 #8230 #10003 #x2610 #x2611 #x2612 #x2613?><?Pub Inc?><kh:document """)#<kh:document xmlns:core="http://www.lexisnexis.com/namespace/sslrp/core" xmlns:fn="http://www.lexisnexis.com/namespace/sslrp/fn" xmlns:header="http://www.lexisnexis.com/namespace/uk/header" xmlns:kh="http://www.lexisnexis.com/namespace/uk/kh" xmlns:lnb="http://www.lexisnexis.com/namespace/uk/lnb" xmlns:lnci="http://www.lexisnexis.com/namespace/common/lnci" xmlns:tr="http://www.lexisnexis.com/namespace/sslrp/tr">""")
    newdata = newdata.replace("[PA]", constantPA)
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



    #wait = input("PAUSED...when ready press enter")

def FindMostRecentFile(directory, pattern):
    try:
        filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
        filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
        return filelist[0]
    except: return 'na'


#main script
print("XML auto-generation for highlights...\n\n")

#Directories
reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
#reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
#reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'
#outputDir = 'xml\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'


#wait = input("PAUSED...when ready press enter")


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
