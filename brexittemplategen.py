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

def FindNextWeekday(givendate, weekday):
    givendate += datetime.timedelta(days=1)
    dayshift = (weekday - givendate.weekday()) % 7
    return givendate + datetime.timedelta(days=dayshift)

    
def TemplateGeneration(PA, highlightdate, highlightType, outputDir, NSMAP):    
    constantPA = PA        
    #PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*' + constantPA + ' Weekly highlights *.xml')
    PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*.xml')
    print('lasthighlightsfilepath: ' + PrevHighlightsFilepath)

    #Extract sections from other docs
    NewsAlertSection = HarvestTemplateSection(templateFilepath, 'Daily and weekly news alerts', NSMAP)
    HeadlinesSection = HarvestTemplateSection(PublicLawFilepath, 'Brexit headlines', NSMAP)
    LegislationSection = HarvestTemplateSection(PublicLawFilepath, 'Brexit legislation', NSMAP)
    SIsSection = HarvestTemplateSection(PublicLawFilepath, 'Brexit SIs', NSMAP)
    LinksSection = HarvestTemplateSection(BrexitTemplateFilepath, 'Brexit content and quick links', NSMAP)
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
        print('Problem loading previous xml for: ' + constantPA)


    #if PA == 'Life Sciences': PA = 'Life Sciences and Pharmaceuticals'
    #if PA == 'Share Incentives': PA = 'Share Schemes'
    #if PA == 'Insurance and Reinsurance': PA = 'Insurance'
    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
    khbody = etree.SubElement(khdoc, '{%s}body' % NSMAP['kh'])
    
    khdoctitle = etree.SubElement(khbody, '{%s}document-title' % NSMAP['kh'])
    khdoctitle.text = PA + ' [weekly/monthly] highlights—[dd Month yyyy]'
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
    try: khbody.append(HeadlinesSection)
    except: print('No Brexit headlines section found in Public Law...')

    #Legislation section from Public Law highlights
    try: khbody.append(LegislationSection)
    except: print('No Brexit legislation section found in Public Law...')

    #SIs section from Public Law highlights
    try: khbody.append(SIsSection)
    except: print('No Brexit SIs section found in Public Law...')

    #Editor's picks
    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
    coretitle.text = "Editor's picks—the practice area/sector view"    
    corepara = etree.SubElement(trsecmain, '{%s}para' % NSMAP['core'])
    corepara.text = 'This section contains key Brexit news hand-picked by LexisPSL lawyers from their own practice areas.'

    for PA in AllPAs:        
        #print(PA)
        PAFilepath = FindMostRecentFile(outputDir + PA + '\\', '*.xml')
        PABrexitSection = HarvestTemplateSection(PAFilepath, 'Brexit', NSMAP)
        if PABrexitSection != None: 
            trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
            coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
            coretitle.text = PA
            try: 
                trsecsub1.append(PABrexitSection)
                print('Brexit section found in ' + PA)
            except: print('No Brexit section found in ' + PA)
            corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
            corepara.text = 'For further updates from ' + PA + ', see: '
            lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
            lncicite.set('normcite', HighlightsOverviewDict.get(PA))
            lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
            lncicontent.text = PA + ' weekly highlights—Overview'

        
    #Links section  
    try: khbody.append(LinksSection)
    except: print('No links section found...')

    #Brexit QAs
    weeklyNewReportFilepath = FindMostRecentFile(reportDir, '*AICER*_UKPSL_weekly_HL_new_[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9].csv')
    dfNewHighlights = DFCleanup(pd.read_csv(weeklyNewReportFilepath), ['SubtopicShortcut', 'Shortcut', 'SubtopicShortcutOfShortcut'], reportDir + 'new')
    #dfNewHighlights = pd.read_csv(weeklyNewReportFilepath)
    dfNew = dfNewHighlights[dfNewHighlights.ContentItemType == 'QandAs'] 
    dfNew = dfNew[dfNew["DocTitle"].str.contains("Brexit", case=False)]
    dfNew = dfNew.sort_values(['DocTitle'], ascending = True)
    newHighlightCount = len(dfNew)
    print(newHighlightCount, 'new QAs')     
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
            if UnderReview == True:
                comment = etree.Comment('Doc ID: ' + str(DocID) + ' UNDER REVIEW: ' + str(UnderReview))
                corepara.append(comment) #add comment after link

        
    #Useful information
    try: khbody.append(UsefulInfoSection)
    except: print('No Useful information section found...')


    tree = etree.ElementTree(khdoc)
    #xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' Weekly highlights template ' + highlightDate + ' test.xml'
    xmlfilepath = localDir + constantPA + '\\' + constantPA + ' Weekly highlights template ' + highlightDate + ' test.xml'
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
            ContentItemType = df.OriginalContentItemType.iloc[i]
            DocID = df.OriginalContentItemId.iloc[i]
            OriginalContentItemPA = df.OriginalContentItemPA.iloc[i]

        list1 = [[DocID, ContentItemType, PA, OriginalContentItemPA, DocTitle, UnderReview]]            
        df1 = df1.append(list1) #append list to dataframe, export to csv outside of the loop
        #print(list1)
        i=i+1 #increment the counter   
    df1.columns = ["DocID", "ContentItemType", "PA", "OriginalContentItemPA", "DocTitle", "UnderReview"]   
    #df1.to_csv(ReportType + '.csv', sep=',',index=False) 
    return df1

#main script
print("Template auto-generation for highlights...\n")
templateFilepath = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Templates\\All Highlights Template.xml'
localDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\xml\\Practice Areas\\'
#outputDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\xml\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'
reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
#reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
#reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
PublicLawFilepath = FindMostRecentFile(outputDir + 'Public Law\\', '*.xml')
BrexitTemplateFilepath = ('\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\Brexit\\Brexit highlights template.xml')
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'


print('Most recent Public Law highlight doc to harvest: \n' + PublicLawFilepath)

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
HighlightsOverviewDict = {"Arbitration": "urn:editpractgovw:8073C915-D589-429E-85E5-F6C8005B41C4", "Banking and Finance": "urn:editpractgovw:58DD7046-895E-46D9-BCE7-670752E807A2", "Commercial": "urn:editpractgovw:5C878F6D-BFE5-44B0-902B-ABBB0DE86071", "Competition": "urn:editpractgovw:C21A4E17-C8D8-4406-B831-21B9DF8B3FD4", "Construction": "urn:editpractgovw:801E3A8B-6F14-47DF-ADD4-9CA3A68E7132", "Corporate": "urn:editpractgovw:0A866544-D60E-401F-9AC5-3F557EB72EA5", "Corporate Crime": "urn:editpractgovw:66046005-2000-4095-95BD-1F9B5273936D", "Dispute Resolution": "urn:editpractgovw:CF409E6B-7930-4510-89FA-8EECFD824E1B", "Employment": "urn:editpractgovw:B698303D-8869-4970-BE78-42B7BBB2CFFE", "Energy": "urn:editpractgovw:906E991B-3579-4780-BCE9-E1ECEB16F2AD", "Environment": "urn:editpractgovw:8E76C46D-B1C8-47B2-B40D-81D3427C125E", "Family": "urn:editpractgovw:D8398746-85F0-4241-AD1A-BE036054F0ED", "Financial Services": "urn:editpractgovw:6F07D371-A265-4983-98A5-A7EDE56957D7", "Immigration": "urn:editpractgovw:4B45F899-A278-4004-9F83-08D87207E3C1", "Information Law": "urn:editpractgovw:613B3C9A-F1E8-42DD-A98B-AA030E002D84", "Insurance": "urn:editpractgovw:F5F3A6FE-8BBA-418B-86DD-6ECC167B1360", "IP": "urn:editpractgovw:E3D61B89-76D9-4676-B15C-A0612FD8A0A7", "Life Sciences and Pharmaceuticals": "urn:editpractgovw:F72AEE92-5B93-48C1-885B-3E40A01C36C8", "Local Government": "urn:editpractgovw:FA30E9E1-C3B0-443C-9B0D-1F8BBD96532C", "Pensions": "urn:editpractgovw:4DA425C3-ECD3-4859-94B8-0A1E32350336", "Personal Injury": "urn:editpractgovw:8C5A7268-F99E-40AE-A4C1-80730567097E", "Planning": "urn:editpractgovw:6D28D29F-63A0-4E45-B556-23BF2AB6CE16", "Practice Compliance": "urn:editpractgovw:14761791-5F12-4917-A764-61C2EFC932EE", "Private Client": "urn:editpractgovw:93561073-3703-4F33-89D1-2CFE104DA262", "Property": "urn:editpractgovw:C3406B32-888E-40D1-BC99-DEBEBA7C39A4", "Property Disputes": "urn:editpractgovw:5FC2E55E-9616-4EEE-823C-26A48403BE9F", "Public Law": "urn:editpractgovw:6D97698D-FB93-484A-A5E5-1896CF796E10", "Restructuring and Insolvency": "urn:editpractgovw:9A232D5A-EA63-4B8F-A2C9-C4702D2EFAB4", "Risk and Compliance": "urn:editpractgovw:2FD8C53E-8960-420B-936A-E1B52C1FF521", "Share Schemes": "urn:editpractgovw:061439F5-C1E5-4D51-968B-1D3277CDB2E8", "Tax": "urn:editpractgovw:B4F236F0-7D46-4568-8D80-E88FDEA1424A", "TMT": "urn:editpractgovw:D0CDDF42-5598-48D9-985C-0F895D10AD50"}
NSMAP = {'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}
highlightType = 'weekly'
givendate = datetime.datetime.today()
#givendate = datetime.date(2019, 12, 31)
nextThursday = FindNextWeekday(givendate, 3) # 3 is thursday


print('Generating Brexit weekly templates for the coming Thursday: ', str(nextThursday.strftime("%#d %B %Y"))) #the hash character turns off the leading zero in the day
PA = 'Brexit'
highlightDate = str(nextThursday.strftime("%#d %B %Y"))
TemplateGeneration(PA, highlightDate, highlightType, outputDir, NSMAP)

#wait = input("PAUSED...when ready press enter")
print('Finished')
