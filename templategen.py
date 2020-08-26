#Template generator for the weekly and monthly highlights
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
import shutil

def FindNextWeekday(givendate, weekday):
    givendate += datetime.timedelta(days=1)
    dayshift = (weekday - givendate.weekday()) % 7
    return givendate + datetime.timedelta(days=dayshift)

def FindLastWorkingDayOfMonth(givendate):
    offset = BMonthEnd()
    givendate += datetime.timedelta(days=3)
    lastday = offset.rollforward(givendate)
    return lastday

def FindLastFridayOfMonth(givendate, weekday):
    lastday = max(week[-3] for week in calendar.monthcalendar(givendate.year, givendate.month))
    if givendate.day == lastday:
        givendate += datetime.timedelta(days=7)
        lastday = max(week[-3] for week in calendar.monthcalendar(givendate.year, givendate.month))
    
    lastday = datetime.date(givendate.year, givendate.month, lastday)
    return lastday

    #dayshift = (weekday - givendate.weekday()) % 7
    #return givendate + datetime.timedelta(days=dayshift)
    
    
def template_generation(PA, highlightdate, highlightType, outputDir, NewsAlertSection, LexTalkSection, NSMAP):  
    log(PA)
    constantPA = PA     
    PA_dir = outputDir + constantPA + '\\'
    list_of_existing_templates = ListOfFilesInDirectory(PA_dir, '*template.xml')
    
    archive_details = Archive(list_of_existing_templates, PA_dir)    
    log('Old templates archived: \n' + str(archive_details) + '\n') 

    #PrevHighlightsFilepath = FindMostRecentFile(PA_dir, '*' + constantPA + ' Weekly highlights *.xml')
    PrevHighlightsFilepath = FindMostRecentFile(PA_dir, '*preview.xml')
    if PrevHighlightsFilepath == 'na': #if can't find from manual download into PA folders, go look in the archive
        log('Finding last week highlight doc...')
        PrevHighlightsFilepath = find_last_week_highlight_doc(highlightsArchiveDir, PA)
    log('lasthighlightsfilepath: ' + PrevHighlightsFilepath)

    #extract info from last highlights' doc
    if PA == 'Life Sciences and Pharmaceuticals': PA = 'Life Sciences'
    if PA == 'Restructuring and Insolvency': PA = 'Restructuring & Insolvency'
    #if PA == 'Insurance and Reinsurance': PA = 'Insurance'
        
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

    except: log('Problem loading previous xml for: ' + constantPA)


    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
    khbody = etree.SubElement(khdoc, '{%s}body' % NSMAP['kh'])
    
    khdoctitle = etree.SubElement(khbody, '{%s}document-title' % NSMAP['kh'])
    khdoctitle.text = PA + ' [weekly/monthly] highlights—[dd Month yyyy]'
    khminisummary = etree.SubElement(khbody, '{%s}mini-summary' % NSMAP['kh'])
    khminisummary.text = "This [week's/month's] edition of [PA] [weekly/monthly] highlights includes:"

    #News
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = '[Topic provided]'
    trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
    coretitle.text = '[News analysis name]'
    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
    corepara.text = '[Mini-summary]'
    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
    corepara.text = 'See News Analysis: [XML ref for News Analysis].'
        
        
    #Covid2019Section news alerts    
    try: khbody.append(Covid2019Section)
    except: log('No Covid2019Section section found...')

        
    #LexTalkSection news alerts    
    try: khbody.append(LexTalkSection)
    except: log('No LexTalkSection section found...')

        
    #Daily and weekly news alerts    
    try: khbody.append(NewsAlertSection)
    except: log('No Daily and weekly news alerts section found...')

    #New and updated content
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'New and updated content'

    #Dates for your diary    
    try: khbody.append(DatesSection)
    except: log('No Dates for your diary section found...')
    
    #Trackers    
    try: khbody.append(TrackersSection)
    except: log('No Tracker section found...')  

    #QAs
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = 'New Q&As'

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
            print('ERROR getting dpsi and doc ID from:')
            print(normcite.get('normcite'))
            continue #lnci cite has no normcite value, so skip as there's nothing to replace
        
        pguid = get_pguid_from_dpsi_docid(dpsi, docid)
        log('Replacing ' + normcite.get('normcite') + ' for ' + pguid)
        normcite.set('normcite', pguid)        



    tree = etree.ElementTree(khdoc)
    if highlightType == 'weekly':
        #REMOVE ONCE ALL DOING WEEKLY
        if PA in MonthlyPAs:
            xmlfilepath = outputDir + constantPA + '\\weekly\\' + PA + ' Weekly highlights ' + highlightDate + ' template.xml'  
        else:
            xmlfilepath = PA_dir + PA + ' Weekly highlights ' + highlightDate + ' template.xml'  
        #xmlfilepath = outputDir + constantPA + '\\Auto-generate XML\\' + PA + ' Weekly highlights ' + highlightDate + '.xml'
    else:
        xmlfilepath = PA_dir + PA + ' Monthly highlights ' + highlightDate + ' template.xml'
        #xmlfilepath = outputDir + constantPA + '\\Auto-generate XML\\' + PA + ' Monthly highlights ' + highlightDate + '.xml'
    tree.write(xmlfilepath, encoding="utf-8", xml_declaration=True, doctype=doctype)

    f = open(xmlfilepath,'r', encoding='utf-8')
    filedata = f.read()
    f.close()
    newdata = filedata  
    PA = PA.replace(" & ", ' &amp; ')
    newdata = newdata.replace(" & ", ' &amp; ')
    newdata = newdata.replace("[PA]", PA)
    newdata = newdata.replace("[weekly/monthly]", highlightType)
    newdata = newdata.replace("[dd Month yyyy]", highlightDate)
    if highlightType == 'weekly':
        newdata = newdata.replace("[week's/month's]", "week's")
    if highlightType == 'monthly':
        newdata = newdata.replace("[week's/month's]", "month's")
    f = open(xmlfilepath,'w', encoding='utf-8')
    f.write(newdata)
    f.close()

    log('XML exported to...' + xmlfilepath + '\n')


def FindMostRecentFile(directory, pattern):
    try:
        filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
        filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
        return filelist[0]
    except: return 'na'

def find_last_week_highlight_doc(directory, PA):    
    filelist = glob.iglob(os.path.join(directory, '*.xml')) #builds list of file in a directory based on a pattern
    for filepath in filelist:
        if filepath.find('0S4D.xml') == -1:
            #print(filepath)
            tree = etree.parse(filepath)
            root = tree.getroot()
            metaData = root.xpath("//header:metadata-item[@name='master-topic-link-parameters-3']", namespaces=NSMAP)
            if len(metaData) == 0: metaData = root.xpath("//header:metadata-item[@name='topic-link-parameters-3']", namespaces=NSMAP)

            #print(etree.tostring(metaData[0]))
            try: metaDataValue = re.search('^::([^:]*):',metaData[0].get('value')).group(1)
            except AttributeError: metaDataValue = re.search('^([^:]*):',metaData[0].get('value')).group(1)
            metaDataValue = re.sub(' $', '', metaDataValue) #remove whitespace at the end of the string if present
            if metaDataValue == 'IP and IT': metaDataValue = 'IP'
            if metaDataValue == 'Share Incentives': metaDataValue = 'Share Schemes'
            if metaDataValue == 'InHouse Advisor': metaDataValue = 'In-House Advisor'
            if metaDataValue == 'Life Sciences': metaDataValue = 'Life Sciences and Pharmaceuticals'
            #print(metaDataValue, PA)
            if metaDataValue == PA: 
                #print('Match found returning filepath: ' + filepath)
                return filepath        
    return 'na'
    

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
        log('Problem loading template xml...')
        return 'section not found'

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
        directory = re.search('(.*\\\\)[^\.]*\.xml',existingFilepath).group(1)
        filename = re.search('.*\\\\([^\.]*\.xml)',existingFilepath).group(1)
        destinationFilepath = directory + 'Archive\\' + filename
        shutil.copy(existingFilepath, destinationFilepath) #Copy
        os.remove(existingFilepath) #Delete old file
        copyList.append('Moved: ' + existingFilepath + ', to: ' + destinationFilepath)
    return copyList


def IsLastWorkingDayOfMonth(givendate):
    offset = BMonthEnd()
    #givendate += datetime.timedelta(days=3)
    lastday = offset.rollforward(givendate)
    if givendate == lastday.date(): return True
    else: return False 

def IsThursday(givendate):
    #print(givendate.weekday())
    if givendate.weekday() == 3: return True
    else: return False


def log(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()
    print(message)


#givendate is auto generated by datetime.datetime.today(), but if you need to manually override it you can set it with datetime.date(year, month, day)


pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'
templateFilepath = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Templates\\All Highlights Template.xml'
highlightsArchiveDir = '\\\\atlas\\Knowhow\\HighlightsArchive\\'

#Updated 150420 1756 - added archiving feature, now archives old templates, also improved the logging
#Updated 170420 1013 - fixed xml metadata find and regex issue in highlights archive docs that have slightly different metadata
#Updated 240420 1735 - converted ampersands to html code, causing issues downstream
#Updated 290420 2229 - convert old dpsi doc id links to pguids with a lookup

state = 'live'
#state = 'livedev'

if state == 'live':
    logDir = "\\\\atlas\\lexispsl\\Highlights\\logs\\"
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
    givendate = datetime.datetime.today()
    givenstrdate =  str(time.strftime("%Y-%m-%d"))
    #givendate = datetime.date(2020, 4, 23)
    #givenstrdate = "2020-04-23"
if state == 'livedev': 
    logDir = "\\\\atlas\\lexispsl\\Highlights\\dev\\logs\\"
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'
    givendate = datetime.date(2020, 4, 23)
    #givendate = datetime.datetime.today()
    givenstrdate = "2020-04-23"

NSMAP = {'lnci': "http://www.lexisnexis.com/namespace/common/lnci", 'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}
AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
#AllPAs = ['Restructuring and Insolvency']
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
MonthlyDates = ['2019-11-29', '2019-12-31', '2020-01-31', '2020-02-28', '2020-03-31', '2020-04-29', '2020-05-29', '2020-06-30', '2020-07-31', '2020-08-28', '2020-09-30', '2020-10-30', '2020-11-30', '2020-12-18']
doctype = '<!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd">'


JCSLogFile = logDir + 'JCSlog-templategen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()

#main script
start = datetime.datetime.now() 
log("Template auto-generation for highlights...\n")
log("Today's date is: " + str(givendate))
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

nextThursday = FindNextWeekday(givendate, 3) # 3 is thursday
lastWorkingDayOfMonth = FindLastWorkingDayOfMonth(givendate)
NewsAlertSection = HarvestTemplateSection(templateFilepath, 'Daily and weekly news alerts', NSMAP)
Covid2019Section = HarvestTemplateSection(templateFilepath, 'Coronavirus', NSMAP)
LexTalkSection = HarvestTemplateSection(templateFilepath, 'LexTalk', NSMAP)

if IsThursday(givendate) == True:
    highlightDate = str(nextThursday.strftime("%#d %B %Y"))
    log('Generating weekly templates for the coming Thursday: ' + highlightDate)
    for PA in AllPAs:    
        #if PA not in MonthlyPAs:
        template_generation(PA, highlightDate, 'weekly', outputDir, NewsAlertSection, LexTalkSection, NSMAP)

else: log('Today is not a Thursday...skipping weekly highlight TEMPLATE generation...')

if str(givenstrdate) in MonthlyDates:
    highlightDate = str(lastWorkingDayOfMonth.strftime("%B %Y"))
    log('Generating monthly templates for the last working day of the month: ' + highlightDate)
    for PA in MonthlyPAs:
        template_generation(PA, highlightDate, 'monthly', outputDir, NewsAlertSection, LexTalkSection, NSMAP)
else: log('Today is not the last working day of the month...skipping monthly highlight TEMPLATE generation...')

    #wait = input("PAUSED...when ready press enter")

log("Finished! Time taken..." + str(datetime.datetime.now() - start))
