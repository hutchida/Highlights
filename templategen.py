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
    
    
def TemplateGeneration(PA, highlightdate, highlightType, outputDir, NewsAlertSection, NSMAP):    
    constantPA = PA        
    #PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*' + constantPA + ' Weekly highlights *.xml')
    PrevHighlightsFilepath = FindMostRecentFile(outputDir + constantPA + '\\', '*.xml')
    print('lasthighlightsfilepath: ' + PrevHighlightsFilepath)
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
        
    #Daily and weekly news alerts    
    try: khbody.append(NewsAlertSection)
    except: print('No Daily and weekly news alerts section found...')
        
    #Dates for your diary    
    try: khbody.append(DatesSection)
    except: print('No Dates for your diary section found...')
    
    #Trackers    
    try: khbody.append(TrackersSection)
    except: print('No Tracker section found...')    
        
    #Useful information
    try: khbody.append(UsefulInfoSection)
    except: print('No Useful information section found...')


    tree = etree.ElementTree(khdoc)
    if constantPA not in MonthlyPAs:
        xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' Weekly highlights ' + highlightDate + ' test.xml'
    else:
        xmlfilepath = outputDir + constantPA + '\\' + constantPA + ' Monthly highlights ' + highlightDate + ' test.xml'
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

#main script
print("Template auto-generation for highlights...\n\n")
templateFilepath = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Templates\\All Highlights Template.xml'
#outputDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\xml\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'

NSMAP = {'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}
    

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance and Reinsurance', 'IP', 'Life Sciences', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Incentives', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance and Reinsurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    

try:
    highlightType = sys.argv[1] #taken from command line
except:
    highlightType = 'weekly'

givendate = datetime.datetime.today()
#givendate = datetime.date(2019, 12, 31)
nextThursday = FindNextWeekday(givendate, 3) # 3 is thursday
lastWorkingDayOfMonth = FindLastWorkingDayOfMonth(givendate)

print('Last working day of the month is: ', str(lastWorkingDayOfMonth.strftime("%#d %B %Y")))
print('Coming Thursday is: ', str(nextThursday.strftime("%#d %B %Y"))) #the hash character turns off the leading zero in the day

NewsAlertSection = HarvestTemplateSection(templateFilepath, 'Daily and weekly news alerts', NSMAP)
#print(NewsAlertSection)

for PA in AllPAs:    
    if PA not in MonthlyPAs:
        highlightDate = str(nextThursday.strftime("%#d %B %Y"))
    else:
        highlightDate = str(lastWorkingDayOfMonth.strftime("%B %Y"))
        TemplateGeneration(PA, highlightDate, highlightType, outputDir, NewsAlertSection, NSMAP)

    #TemplateGeneration(PA, highlightDate, highlightType, outputDir, NewsAlertSection, NSMAP)
    #wait = input("PAUSED...when ready press enter")
print('Finished')
