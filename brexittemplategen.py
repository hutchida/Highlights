#Brexit template generator for the weekly highlights
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
import calendar
import sys
import xml.etree.ElementTree as ET
from lxml import etree
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import shutil

def FindNextWeekday(givendate, weekday):
    givendate += datetime.timedelta(days=1)
    dayshift = (weekday - givendate.weekday()) % 7
    return givendate + datetime.timedelta(days=dayshift)


def FirstWeekOfTheMonth():
    today = datetime.datetime.today()
    if today.weekday() - today.day >= -1: return True
    else: return False

def FindLastWeekHighLightDoc(directory, PA):    
    filelist = glob.iglob(os.path.join(directory, '*.xml')) #builds list of file in a directory based on a pattern
    for filepath in filelist:        
        if filepath.find('0S4D.xml') == -1:
            #print(filepath)
            tree = etree.parse(filepath)
            root = tree.getroot()
            metaData = root.xpath("//header:metadata-item[@name='master-topic-link-parameters-3']", namespaces=NSMAP)        
            if len(metaData) == 0: metaData = root.xpath("//header:metadata-item[@name='topic-link-parameters-3']", namespaces=NSMAP)
            
            try: metaDataValue = re.search('^::([^:]*):',metaData[0].get('value')).group(1)
            except AttributeError: metaDataValue = re.search('^([^:]*):',metaData[0].get('value')).group(1)

            metaDataValue = re.sub(' $', '', metaDataValue) #remove whitespace at the end of the string if present
            if metaDataValue == 'IP and IT': metaDataValue = 'IP'
            if metaDataValue == 'Share Incentives': metaDataValue = 'Share Schemes'
            if metaDataValue == 'InHouse Advisor': metaDataValue = 'In-House Advisor'
            if metaDataValue == 'Life Sciences': metaDataValue = 'Life Sciences and Pharmaceuticals'
            #print(metaDataValue, PA)
            if metaDataValue == PA: 
                print('Match found returning filepath: ' + filepath)
                return filepath        
    return 'na'

def template_generation(PA, highlightdate, highlightType, outputDir, NSMAP):    
    constantPA = PA        
    #PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*' + constantPA + ' Weekly highlights *.xml')
    PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*preview.xml')
    log('lasthighlightsfilepath: ' + PrevHighlightsFilepath)

    #Extract sections from other docs
    NewsAlertSection = HarvestTemplateSection(templateFilepath, 'Daily and weekly news alerts', NSMAP)
    HeadlinesSection = HarvestTemplateSection(PublicLawFilepath, 'Brexit headlines', NSMAP)
    LegislationSection = HarvestTemplateSection(PublicLawFilepath, 'Brexit legislation', NSMAP)
    SIsSection = HarvestTemplateSection(PublicLawFilepath, 'Brexit SIs', NSMAP)
    LinksSection = HarvestTemplateSection(BrexitTemplateFilepath, 'Brexit content and quick links', NSMAP)
    LexTalkSection = HarvestTemplateSection(templateFilepath, 'LexTalk', NSMAP)
    UsefulInfoSection = HarvestTemplateSection(BrexitTemplateFilepath, 'Useful information', NSMAP)

    #extract info from last highlights' doc    
    try:
        tree = etree.parse(PrevHighlightsFilepath)
        root = tree.getroot()

        trsecmains = root.findall('.//tr:secmain', NSMAP)
        for trsecmain in trsecmains:
            coretitle = trsecmain.find('core:title', NSMAP)
            try:
                del coretitle.attrib["id"]
                #print('\nDELETED id attribute...' + coretitle.text)
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
        log('Problem loading previous xml for: ' + constantPA)


    #if PA == 'Life Sciences': PA = 'Life Sciences and Pharmaceuticals'
    #if PA == 'Share Incentives': PA = 'Share Schemes'
    #if PA == 'Insurance and Reinsurance': PA = 'Insurance'
    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
    khbody = etree.SubElement(khdoc, '{%s}body' % NSMAP['kh'])
    
    khdoctitle = etree.SubElement(khbody, '{%s}document-title' % NSMAP['kh'])
    khdoctitle.text = PA + ' highlights—[dd Month yyyy]'
    khminisummary = etree.SubElement(khbody, '{%s}mini-summary' % NSMAP['kh'])
    khminisummary.text = "These Brexit highlights bring you a summary of the latest Brexit news and legislation updates from across a range of LexisNexis® practice areas."
    #Intro
    corepara = etree.SubElement(khbody, '{%s}para' % NSMAP['core'])
    corepara.text = 'For guidance on keeping up to date, including details of how to access the latest Brexit news updates and analysis, plus instructions for setting up daily/weekly alerts via email and RSS, see: '
    lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
    lncicite.set('normcite', 'urn:editpractgfaq:A0F1E667-F841-42F6-9657-00BCA2A4AB03')
    lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
    lncicontent.text = 'How do I sign up for Brexit alerts?'

    #Headline section from Public Law highlights
    try: 
        Headlines = HeadlinesSection.findall('.//tr:secsub1', namespaces=NSMAP)  
        trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
        coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
        coretitle.text = 'General Brexit headlines'
        corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
        corepara.text = 'This section contains key overarching Brexit news headlines.'
        for Headline in Headlines:
            HeadlineTitle = Headline.find('{%s}title' % NSMAP['core'])
            HeadlineParas = Headline.findall('{%s}para' % NSMAP['core'])
            trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
            coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
            coretitle.text = HeadlineTitle.text
            for HeadlinePara in HeadlineParas:                
                #corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                #corepara.text = HeadlinePara.text
                trsecsub1.append(HeadlinePara) 
    except: 
        log('No Brexit headlines section found in Public Law...')



    #Legislation section from Public Law highlights
    try: 
        Legs = LegislationSection.findall('.//tr:secsub1', namespaces=NSMAP)  
        trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
        coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
        coretitle.text = 'Brexit legislation updates'
        corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
        corepara.text = 'This section contains Brexit news headlines relating to Brexit-related primary legislation and legislative preparation for Brexit generally.'
        for Leg in Legs:
            LegTitle = Leg.find('{%s}title' % NSMAP['core'])
            trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
            coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
            coretitle.text = LegTitle.text
            LegParas = Leg.findall('{%s}para' % NSMAP['core'])
            for LegPara in LegParas:
                #corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                #corepara.text = LegPara.text
                trsecsub1.append(LegPara) 
    except: 
        log('No Brexit legislation section found in Public Law...')


    #SIs section from Public Law highlights
    try: 
        SIs = SIsSection.findall('.//tr:secsub1', namespaces=NSMAP)  
        trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
        coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
        coretitle.text = 'Brexit SIs and sifting updates'
        corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
        corepara.text = 'This section contains updates on the latest final and draft Brexit SIs laid in Parliament, plus updates on proposed negative Brexit SIs laid for sifting.'
        for SI in SIs:
            SITitle = SI.find('{%s}title' % NSMAP['core'])            
            trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
            coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
            coretitle.text = SITitle.text
            SIParas = SI.findall('{%s}para' % NSMAP['core'])
            for SIPara in SIParas:
                #corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                #corepara.text = SIPara.text
                trsecsub1.append(SIPara) 
    except: 
        log('No Brexit SIs section found in Public Law...')

    #Made Brexit SIs laid in Parliament
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'Made Brexit SIs laid in Parliament'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'xxx'

    #Draft Brexit SIs laid in Parliament
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'Draft Brexit SIs laid in Parliament'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'xxx'

    #Draft Brexit SIs laid for sifting and sifting committee recommendations
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'Draft Brexit SIs laid for sifting and sifting committee recommendations'
    #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    #corepara.text = 'xxx'

    #Editor's picks
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = "Editor's picks—the practice area/sector view"    
    corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    corepara.text = 'This section contains key Brexit news hand-picked by Lexis®PSL lawyers from their own practice areas.'
    log('Harvesting Brexit sections from Editors Picks in other PAs weekly highlights from...')
    for PA in AllPAs:        
        log(PA)
        
        PAFilepath = FindLastWeekHighLightDoc(highlightsArchiveDir, PA)
        if PAFilepath =='na': PAFilepath = FindMostRecentFile(outputDir + PA + '\\', '*preview.xml')
            
        try: 
            PABrexitSection = HarvestTemplateSection(PAFilepath, 'Brexit', NSMAP) 
            PABrexitSectionExists = True
        except:
            PABrexitSectionExists = False

             
        if PABrexitSectionExists == True:   
            log(PAFilepath)    
            if PA != 'Public Law':
                if FirstWeekOfTheMonth() == True:
                    if PABrexitSection != None: 
                        PASecSub1 = PABrexitSection.find('.//tr:secsub1', namespaces=NSMAP)   #find the trsecsub1
                        PASecSub1Title = PASecSub1.find('{%s}title' % NSMAP['core']) #find within it the title
                        if PASecSub1Title != None: #check that there is text in the title tags, else skip
                            if PA == 'Life Sciences and Pharmaceuticals': PA = 'Life Sciences'
                            trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
                            coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
                            coretitle.text = PA
                            
                            log('Brexit section found in ' + PA)
                            #print(PABrexitSection)
                            PASecSub1s = PABrexitSection.findall('.//tr:secsub1', namespaces=NSMAP)            
                            for PASecSub1 in PASecSub1s:
                                PASecSub1Title = PASecSub1.find('{%s}title' % NSMAP['core'])
                                #print(PASecSub1Title.text)
                                
                                corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                coreparabold = etree.SubElement(corepara, '{%s}emph' % NSMAP['core'])                    
                                coreparabold.set('typestyle', 'bf')
                                coreparabold.text = PASecSub1Title.text    

                                PASecSub1Paras = PASecSub1.findall('{%s}para' % NSMAP['core'])
                                for PASecSub1Para in PASecSub1Paras:
                                    try:
                                        trsecsub1.append(PASecSub1Para) 
                                    except:
                                        corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                    
                                    #corepara.text = PASecSub1Para.text
                            corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                            corepara.text = 'For further updates from ' + PA + ', see: '
                            lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                            lncicite.set('normcite', HighlightsOverviewDict.get(PA))
                            lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                            lncicontent.text = PA + ' weekly highlights—overview'
                            lncicite.tail = '.'
                        else:
                            comment = etree.Comment(PA + ": no Brexit news found")
                            trsecmain.append(comment) 
                    else:
                        comment = etree.Comment(PA + ": no Brexit news found")
                        trsecmain.append(comment) 
                else:
                    if PA not in MonthlyPAs:
                        if PABrexitSection != None: 
                            PASecSub1 = PABrexitSection.find('.//tr:secsub1', namespaces=NSMAP)   #find the trsecsub1
                            if PASecSub1 != None: 
                                PASecSub1Title = PASecSub1.find('{%s}title' % NSMAP['core']) #find within it the title
                                if PASecSub1Title != None: #check that there is text in the title tags, else skip
                                    if PA == 'Life Sciences and Pharmaceuticals': PA = 'Life Sciences'
                                    trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
                                    coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
                                    coretitle.text = PA
                                    
                                    log('Brexit section found in ' + PA)
                                    #print(PABrexitSection)
                                    PASecSub1s = PABrexitSection.findall('.//tr:secsub1', namespaces=NSMAP)            
                                    for PASecSub1 in PASecSub1s:
                                        PASecSub1Title = PASecSub1.find('{%s}title' % NSMAP['core'])
                                        #print(PASecSub1Title.text)
                                        
                                        corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                        coreparabold = etree.SubElement(corepara, '{%s}emph' % NSMAP['core'])                    
                                        coreparabold.set('typestyle', 'bf')
                                        coreparabold.text = PASecSub1Title.text    

                                        PASecSub1Paras = PASecSub1.findall('{%s}para' % NSMAP['core'])
                                        for PASecSub1Para in PASecSub1Paras:
                                            try:
                                                trsecsub1.append(PASecSub1Para) 
                                            except:
                                                corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                            
                                            #corepara.text = PASecSub1Para.text
                                    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                    corepara.text = 'For further updates from ' + PA + ', see: '
                                    lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                                    lncicite.set('normcite', HighlightsOverviewDict.get(PA))
                                    lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                                    lncicontent.text = PA + ' weekly highlights—overview'
                                    lncicite.tail = '.'
                                else:
                                    comment = etree.Comment(PA + ": no Brexit news found")
                                    trsecmain.append(comment) 
                            else:
                                comment = etree.Comment(PA + ": no Brexit news found")
                                trsecmain.append(comment) 
                        else:
                            comment = etree.Comment(PA + ": no Brexit news found")
                            trsecmain.append(comment) 
                    else:
                        log('Not the first week of the month and this PA is a monthly highlight, so skipping the brexit news harvest: ' + PA)
            else:
                trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
                coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
                coretitle.text = PA
                corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                coreparabold = etree.SubElement(corepara, '{%s}emph' % NSMAP['core'])                    
                coreparabold.set('typestyle', 'bf')
                coreparabold.text = '***insert title here***'                  
                corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                corepara.text =  '***insert para here***'
                corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                corepara.text = 'For further updates from ' + PA + ', see: '
                lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                lncicite.set('normcite', HighlightsOverviewDict.get(PA))
                lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                lncicontent.text = PA + ' weekly highlights—overview'
                lncicite.tail = '.'
        else:
            comment = etree.Comment(PA + ": no highlight doc found")
            trsecmain.append(comment) 

        #Archive section - archive all of last week's highlights docs
        #existingDocs = ListOfFilesInDirectory(outputDir + PA + '\\', '*preview.xml')
        #archivedList = Archive(existingDocs, outputDir + PA + '\\')
        #print('Moved following files to Archive folder: ' + str(archivedList))
        #log('Moved following files to Archive folder: ' + str(archivedList))
    #Links section  
    try: khbody.append(LinksSection)
    except: print('No links section found...')

        #BrexitLinkTitle = LinksSection.find('.//core:title', namespaces=NSMAP)       
        #BrexitLinksPara = LinksSection.find('.//core:para', namespaces=NSMAP)   
        #trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
        #coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
        #coretitle.text = BrexitLinkTitle.text
        #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
        #corepara.text = BrexitLinksPara.text
        #corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
            #corelist.set('type', 'bullet')

    BrexitLinks = LinksSection.findall('.//core:listitem', namespaces=NSMAP)       
    for BrexitLink in BrexitLinks:             
        trsecmain.append(BrexitLink)

    #Brexit new or updated docs, i.e. Links section
    weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    log('Latest weeklyNewReportFilepath:' + weeklyNewReportFilepath)
    weeklyUpdateReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_updated_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    log('Latest weeklyUpdateReportFilepath:' + weeklyUpdateReportFilepath)
    dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilepath), ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut'], reportDir + 'new')
    dfUpdateHighlights = DFCleanup(pd.read_csv(weeklyUpdateReportFilepath), ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut'], reportDir + 'updated')
    dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')        


    #Updated and New content
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'New and updated Brexit related content'
    
    #New
    if len(dfNewHighlights[(dfNewHighlights.ContentItemType != 'QandAs') & dfNewHighlights.DocTitle.str.contains("Brexit", case=False)]) > 0: #if there are any new docs for the PA, create new doc section
        
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']
        for ContentType in ContentTypeList:      
            dfNew = dfNewHighlights[dfNewHighlights.ContentItemType != 'QandAs'] 
            dfNew = dfNew[dfNew["DocTitle"].str.contains("Brexit", case=False)]        
            dfNew = dfNew[dfNew.ContentItemType == ContentType]
            #print(len(dfNew), ContentType)
            
            if len(dfNew) > 0: 
                dfNew = dfNew.sort_values(['DocTitle'], ascending = True)
                newHighlightCount = len(dfNew)
                log(ContentType + ' ' + str(newHighlightCount) + ' new docs')
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
                        #print(ContentType, OriginalPA)
                        dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == OriginalPA), 'DPSI'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of DPSI
                        #print (dpsi)
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
    if len(dfUpdateHighlights[(dfUpdateHighlights.ContentItemType != 'QandAs') & dfUpdateHighlights.DocTitle.str.contains("Brexit", case=False)]) > 0: #if there are any new docs for the PA, create new doc section 
        
        
        ContentTypeList=['PracticeNote', 'Precedent', 'Checklist']
        for ContentType in ContentTypeList:    
            dfUpdate = dfUpdateHighlights[dfUpdateHighlights.ContentItemType != 'QandAs'] 
            dfUpdate = dfUpdate[dfUpdate["DocTitle"].str.contains("Brexit", case=False)]
            dfUpdate = dfUpdate[dfUpdate.ContentItemType == ContentType]
            #print(len(dfNew), ContentType)
            if len(dfUpdate) > 0: 
                dfUpdate = dfUpdate.sort_values(['DocTitle'], ascending = True)
                updateHighlightCount = len(dfUpdate)
                #print(ContentType, updateHighlightCount, 'updates') 
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
    
    

    #Brexit QAs
    weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilepath), ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut'], reportDir + 'new')
    #dfNewHighlights = pd.read_csv(weeklyNewReportFilepath)
    dfNew = dfNewHighlights[dfNewHighlights.ContentItemType == 'QandAs'] 
    dfNew = dfNew[dfNew["DocTitle"].str.contains("Brexit", case=False)]
    dfNew = dfNew.sort_values(['DocTitle'], ascending = True)
    newHighlightCount = len(dfNew)
    log(str(newHighlightCount) + ' new QAs')
    if newHighlightCount > 0:     
        dfdpsi = pd.read_csv(lookupdpsi, encoding='utf-8')       
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
                    log(PA)       
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
    else:
        comment = etree.Comment("No Brexit related QAs found")
        khbody.append(comment) 
     
        
    #LexTalkSection news alerts    
    try: khbody.append(LexTalkSection)
    except: 
        log('No LexTalkSection section found...')

        
    #Useful information
    try: khbody.append(UsefulInfoSection)
    except: log('No Useful information section found...')



    def get_pguid_from_dpsi_docid(dpsi, docid):
        try:
            
            #print(DocID, DocTitle, ContentType, dpsi, pguidlookup)
            if dpsi == '0S4D': tag = pguidlookupdom_0S4D.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OLB': tag = pguidlookupdom_0OLB.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OJ8': tag = pguidlookupdom_0OJ8.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OJK': tag = pguidlookupdom_0OJK.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0RU8': tag = pguidlookupdom_0RU8.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OLV': tag = pguidlookupdom_0OLV.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OJW': tag = pguidlookupdom_0OJW.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OM9': tag = pguidlookupdom_0OM9.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0R2W': tag = pguidlookupdom_0R2W.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OLN': tag = pguidlookupdom_0OLN.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0ON9': tag = pguidlookupdom_0ON9.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0ONJ': tag = pguidlookupdom_0ONJ.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0SBX': tag = pguidlookupdom_0SBX.find(".//document[@database-id='" + str(docid) + "']")
            if dpsi == '0OJQ': tag = pguidlookupdom_0OJQ.find(".//document[@database-id='" + str(docid) + "']")
            else:
                pguidlookupdom = etree.parse(pguidlistDir + dpsi + '.xml') 
                tag = pguidlookupdom.find(".//document[@database-id='" + str(docid) + "']")
            return tag.get('pguid')
        except:
            return 'notfound'

    log('Replacing dpsi_docid links for pguids...')
    list_of_lnci_cites = khdoc.findall('.//lnci:cite', NSMAP)
    for normcite in list_of_lnci_cites:
        #print(normcite.get('normcite'))
        try: 
            dpsi = re.search('(.*)_(\d*)',normcite.get('normcite')).group(1)
            docid = re.search('(.*)_(\d*)',normcite.get('normcite')).group(2)
        except: 
            log('ERROR getting dpsi and doc ID from:')
            log(str(normcite.get('normcite')))
            continue #lnci cite has no normcite value, so skip as there's nothing to replace
        
        pguid = get_pguid_from_dpsi_docid(dpsi, docid)
        log('Replacing ' + str(normcite.get('normcite')) + ' for ' + pguid)
        normcite.set('normcite', pguid)        


    tree = etree.ElementTree(khdoc)
    xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' highlights template ' + highlightDate + ' test.xml'
    #xmlfilepath = localDir + constantPA + '\\' + constantPA + ' highlights template ' + highlightDate + ' test.xml'
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

    #Archive the old preview files
    #Archive([PrevHighlightsFilepath], outputDir + constantPA + '\\')
    #print('Archived: ' + PrevHighlightsFilepath)
    #log('Archived: ' + PrevHighlightsFilepath)
    #Archive([PublicLawFilepath], outputDir + constantPA + '\\')
    #print('Archived: ' + PublicLawFilepath)
    #log('Archived: ' + PublicLawFilepath)


    log('XML exported to...' + xmlfilepath)
    #wait = input("PAUSED...when ready press enter")

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

def HarvestTemplateSection(templateFilepath, searchterm, NSMAP):    
    templatetree = etree.parse(templateFilepath)
    templateroot = templatetree.getroot()
    trsecmains = templateroot.findall('.//tr:secmain', NSMAP)
    try:
        for trsecmain in trsecmains:
            coretitle = trsecmain.find('core:title', NSMAP)            
            try:
                if coretitle.text.find(searchterm) > -1: 
                    return trsecmain
            except: pass
    except:
        print('Problem loading template xml...')
        return 'section not found'


def DFCleanup(df, ShortcutTypeList, ReportType):
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
            continue #filter out shortcuts
            #ContentItemType = df.OriginalContentItemType.iloc[i]
            #DocID = df.OriginalContentItemId.iloc[i]
            #OriginalContentItemPA = df.OriginalContentItemPA.iloc[i]

        list1 = [[DocID, ContentItemType, PA, OriginalContentItemPA, DocTitle, UnderReview]]            
        df1 = df1.append(list1) #append list to dataframe, export to csv outside of the loop
        #print(list1)
        i=i+1 #increment the counter   
    df1.columns = ["DocID", "ContentItemType", "PA", "OriginalContentItemPA", "DocTitle", "UnderReview"]   
    #df1.to_csv(ReportType + '.csv', sep=',',index=False) 
    return df1


def log(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()
    print(message)

def formatEmail(Subject, HTML):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = Subject
    msg["From"] = sender_email
    msg["To"] = receiver_email    

    HTMLPart = MIMEText(HTML, "html")
    msg.attach(HTMLPart)
    return msg

def sendEmail(msg, receiver_email):
    s = smtplib.SMTP("LNGWOKEXCP002.legal.regn.net")
    s.sendmail(sender_email, receiver_email, msg.as_string())

#Updated 170420 1013 - fixed xml metadata find and regex issue in highlights archive docs that have slightly different metadata
#200520 1241 - changed to remove restructuring and insolvency from monthly highlight list
#state = 'livedev'
state = 'live'

if state == 'livedev':     
    receiver_email_list = ['daniel.hutchings.1@lexisnexis.co.uk']
    logDir = "\\\\atlas\\lexispsl\\Highlights\\dev\\logs\\"
    templateFilepath = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Templates\\All Highlights Template.xml'
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'
    reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
    
if state == 'live':
    #receiver_email_list = ['daniel.hutchings.1@lexisnexis.co.uk']
    receiver_email_list = ['LNGUKPSLDigitalEditors@ReedElsevier.com', 'holly.nankivell@lexisnexis.co.uk', 'louis.payne@lexisnexis.co.uk', 'anne.kingsley@lexisnexis.co.uk', 'Cristiana.Rossetti@lexisnexis.co.uk', 'michael.agnew@lexisnexis.co.uk', 'james-john.dwyer-wilkinson@lexisnexis.co.uk']
    logDir = "\\\\atlas\\lexispsl\\Highlights\\logs\\"
    templateFilepath = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Templates\\All Highlights Template.xml'
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
    reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Risk and Compliance']    
HighlightsOverviewDict = {"Arbitration": "urn:editpractgovw:8073C915-D589-429E-85E5-F6C8005B41C4", "Banking and Finance": "urn:editpractgovw:58DD7046-895E-46D9-BCE7-670752E807A2", "Commercial": "urn:editpractgovw:5C878F6D-BFE5-44B0-902B-ABBB0DE86071", "Competition": "urn:editpractgovw:C21A4E17-C8D8-4406-B831-21B9DF8B3FD4", "Construction": "urn:editpractgovw:801E3A8B-6F14-47DF-ADD4-9CA3A68E7132", "Corporate": "urn:editpractgovw:0A866544-D60E-401F-9AC5-3F557EB72EA5", "Corporate Crime": "urn:editpractgovw:66046005-2000-4095-95BD-1F9B5273936D", "Dispute Resolution": "urn:editpractgovw:CF409E6B-7930-4510-89FA-8EECFD824E1B", "Employment": "urn:editpractgovw:B698303D-8869-4970-BE78-42B7BBB2CFFE", "Energy": "urn:editpractgovw:906E991B-3579-4780-BCE9-E1ECEB16F2AD", "Environment": "urn:editpractgovw:8E76C46D-B1C8-47B2-B40D-81D3427C125E", "Family": "urn:editpractgovw:D8398746-85F0-4241-AD1A-BE036054F0ED", "Financial Services": "urn:editpractgovw:6F07D371-A265-4983-98A5-A7EDE56957D7", "Immigration": "urn:editpractgovw:4B45F899-A278-4004-9F83-08D87207E3C1", "Information Law": "urn:editpractgovw:613B3C9A-F1E8-42DD-A98B-AA030E002D84", "In-House Advisor": "notfound", "Insurance": "urn:editpractgovw:F5F3A6FE-8BBA-418B-86DD-6ECC167B1360", "IP": "urn:editpractgovw:11AFFA76-B5EF-4733-8A5C-88079D29EC1B", "Life Sciences": "urn:editpractgovw:F72AEE92-5B93-48C1-885B-3E40A01C36C8", "Local Government": "urn:editpractgovw:FA30E9E1-C3B0-443C-9B0D-1F8BBD96532C", "Pensions": "urn:editpractgovw:4DA425C3-ECD3-4859-94B8-0A1E32350336", "Personal Injury": "urn:editpractgovw:8C5A7268-F99E-40AE-A4C1-80730567097E", "Planning": "urn:editpractgovw:6D28D29F-63A0-4E45-B556-23BF2AB6CE16", "Practice Compliance": "urn:editpractgovw:14761791-5F12-4917-A764-61C2EFC932EE",  "Practice Management": "notfound", "Private Client": "urn:editpractgovw:93561073-3703-4F33-89D1-2CFE104DA262", "Property": "urn:editpractgovw:C3406B32-888E-40D1-BC99-DEBEBA7C39A4", "Property Disputes": "urn:editpractgovw:5FC2E55E-9616-4EEE-823C-26A48403BE9F", "Public Law": "urn:editpractgovw:6D97698D-FB93-484A-A5E5-1896CF796E10", "Restructuring and Insolvency": "urn:editpractgovw:9A232D5A-EA63-4B8F-A2C9-C4702D2EFAB4", "Risk and Compliance": "urn:editpractgovw:2FD8C53E-8960-420B-936A-E1B52C1FF521", "Share Schemes": "urn:editpractgovw:061439F5-C1E5-4D51-968B-1D3277CDB2E8", "Tax": "urn:editpractgovw:B4F236F0-7D46-4568-8D80-E88FDEA1424A", "TMT": "urn:editpractgovw:D0CDDF42-5598-48D9-985C-0F895D10AD50", "Wills and Probate": "notfound"}
NSMAP = {'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}

highlightsArchiveDir = '\\\\atlas\\Knowhow\\HighlightsArchive\\'
sender_email = 'LNGUKPSLDigitalEditors@ReedElsevier.com'
PublicLawFilepath = FindLastWeekHighLightDoc(highlightsArchiveDir, 'Public Law')
#PublicLawFilepath = FindMostRecentFile(outputDir + 'Public Law\\', '*preview.xml')
#BrexitTemplateFilepath = ('\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\Brexit\\Brexit highlights template.xml')
BrexitTemplateFilepath = FindMostRecentFile(outputDir + 'Brexit\\', '*preview.xml')
#BrexitTemplateFilepath = FindLastWeekHighLightDoc(highlightsArchiveDir, 'Brexit')
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'

highlightType = 'weekly'
givendate = datetime.datetime.today()
givendate += datetime.timedelta(days=1)

    
JCSLogFile = logDir + 'JCSlog-brexittemplategen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()

#main script
highlightDate = str(givendate.strftime("%#d %B %Y")) #the hash character turns off the leading zero in the day
log('Generating Brexit templates for tomorrow: ' + highlightDate)
PrevTemplateError=0

log('Loading 0S4D.xml and a bunch of other heavy xml files for pguid look up later...')
pguidlookupdom_0S4D = etree.parse(pguidlistDir + '0S4D.xml')
pguidlookupdom_0OLB = etree.parse(pguidlistDir + '0OLB.xml')
pguidlookupdom_0OJ8 = etree.parse(pguidlistDir + '0OJ8.xml')
pguidlookupdom_0OJK = etree.parse(pguidlistDir + '0OJK.xml')
pguidlookupdom_0RU8 = etree.parse(pguidlistDir + '0RU8.xml')
pguidlookupdom_0OLV = etree.parse(pguidlistDir + '0OLV.xml')
pguidlookupdom_0OJW = etree.parse(pguidlistDir + '0OJW.xml')
pguidlookupdom_0OM9 = etree.parse(pguidlistDir + '0OM9.xml')
pguidlookupdom_0R2W = etree.parse(pguidlistDir + '0R2W.xml')
pguidlookupdom_0OLN = etree.parse(pguidlistDir + '0OLN.xml')
pguidlookupdom_0ON9 = etree.parse(pguidlistDir + '0ON9.xml')
pguidlookupdom_0ONJ = etree.parse(pguidlistDir + '0ONJ.xml')
pguidlookupdom_0SBX = etree.parse(pguidlistDir + '0SBX.xml')
pguidlookupdom_0OJQ = etree.parse(pguidlistDir + '0OJQ.xml')

if PublicLawFilepath != 'na':
    log('Most recent Public Law highlight doc to harvest: \n' + PublicLawFilepath)
else:
    log('Problem encountered trying to find the most recent Public Law weekly highlight doc in: ' + outputDir + 'Public Law\\' + ' with the pattern *preview.xml')
    PrevTemplateError+=1
if BrexitTemplateFilepath != 'na':
    log('Most recent Brexit template  to harvest: \n' + BrexitTemplateFilepath)
else:
    log('Problem encountered trying to find the most recent Brexit highlight doc in: ' + outputDir + 'Brexit\\' + ' with the pattern *preview.xml')
    PrevTemplateError+=1

if PrevTemplateError==0:
    PA = 'Brexit'
    template_generation(PA, highlightDate, highlightType, outputDir, NSMAP)

    Subject = "Brexit highlights report - READY"
    HTML = """\
<html>
    <head>
        <title>Brexit highlights report</title>
    </head>
    <body>
        <div style="font-size: 100%; text-align: center;">
            <p>The Brexit highlights template is now ready: <br />
			<a href="file://///atlas/lexispsl/Highlights/Practice Areas/Brexit/" target="_blank">\\\\atlas\lexispsl\Highlights\Practice Areas\Brexit\</a></p>
			<sup>See the log for more details: <br />
			<a href="file://///atlas/lexispsl/Highlights/Automatic creation/Logs/JCSlog-brexittemplategen.txt" target="_blank">\\\\atlas\lexispsl\Highlights\Automatic creation\Logs\JCSlog-brexittemplategen.txt</a></sup>
        </div>
    </body>
</html>
    """

    #create and send email
    for receiver_email in receiver_email_list:
        msg = formatEmail(Subject, HTML)
        sendEmail(msg, receiver_email)
else:
    Subject = "Brexit highlights report - NOT READY"
    HTML = """\
<html>
    <head>
        <title>Brexit highlights report - NOT READY</title>
    </head>
    <body>
        <div style="font-size: 100%; text-align: left;">
            <p>The Brexit Highlights script has not run as last week’s public law and/or Brexit highlight have not been downloaded and saved in their folders by 2.30pm on a Thursday in:</p>
			<p><a href="file://///atlas/lexispsl/Highlights/Practice Areas/Brexit/" target="_blank">\\\\atlas\lexispsl\Highlights\Practice Areas\Brexit\</a><br />
            <a href="file://///atlas/lexispsl/Highlights/Practice Areas/Public Law/" target="_blank">\\\\atlas\lexispsl\Highlights\Practice Areas\Public Law\</a></p>
			<p>or the filename has been amended from the name given by echo (containing '*preview.xml').</p>
            <p>Please download these files or amend the filenames and reply to this email to request the automation script to be rerun.</p>
            <br/><br/>
            <sup>See the log for more details: <br />
			<a href="file://///atlas/lexispsl/Highlights/Automatic creation/Logs/JCSlog-brexittemplategen.txt" target="_blank">\\\\atlas\lexispsl\Highlights\Automatic creation\Logs\JCSlog-brexittemplategen.txt</a></sup>
        </div>
    </body>
</html>
    """
    #create and send email
    for receiver_email in receiver_email_list:
        msg = formatEmail(Subject, HTML)
        sendEmail(msg, receiver_email)
    
    log("Problem with the previous week's highlights docs not being present, email sent to distribution list...")
#wait = input("PAUSED...when ready press enter")
log('Finished')


