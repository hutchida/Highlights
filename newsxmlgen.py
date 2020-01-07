#News XML generator for highlights, loops through all entries on the news items list supplied by Martin/Connell and creates xml with links to those mentioned documents that can be copied and pasted into the highlights templates
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


    
def XMLGeneration(PA, highlightDate, highlightType, outputDir):
    constantPA = PA            
    archiveDir = outputDir + constantPA + '\\archive\\'     
    XLSFilepath = FindMostRecentFile(outputDir + constantPA + '\\Auto-generate XML\\', '*.xlsx')
    if XLSFilepath != 'na':
        try: 
            XLSFilepathDate = re.search('news items (.*)\.xlsx',XLSFilepath).group(1) 
            XLSarchiveFilepath = archiveDir + constantPA + ' news items ' + XLSFilepathDate + '.xlsx'
            XMLFilepath = outputDir + constantPA + '\\' + constantPA + ' news items ' + highlightDate + ' test.xml'
            LogOutput('\nXLSX filepath to check...' + XLSFilepath)
        
            dfPA = pd.ExcelFile(XLSFilepath)
            #LogOutput('Spreadsheet loaded...')
            dfPA = dfPA.parse("Sheet1")
            #LogOutput('Spreadsheet parsed...')
            NSMAP = {'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'lnb': 'http://www.lexisnexis.com/namespace/uk/lnb', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr'}#, 'atict': 'http://www.arbortext.com/namespace/atict'}

            if PA == 'Life Sciences and Pharmaceuticals': PA = 'Life Sciences'
            if PA == 'Banking and Finance': PA = 'Banking &amp; Finance'
            if PA == 'Share Schemes': PA = 'Share Incentives'
            if PA == 'Risk and Compliance': PA = 'Risk &amp; Compliance'
            if PA == 'Wills and Probate': PA = 'Wills &amp; Probate'
            

            newsItemCount = len(dfPA)
            topicList = dfPA['TopicName'].tolist()
            dedupedList = []
            for item in topicList:
                if str(item) != 'nan':
                    if item not in dedupedList: dedupedList.append(item)    
            topicList = dedupedList
            if newsItemCount > 0:          
                khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                khbody = etree.SubElement(khdoc, '{%s}body' % NSMAP['kh'])              
                khdoctitle = etree.SubElement(khbody, '{%s}document-title' % NSMAP['kh'])
                khdoctitle.text = PA + ' [weekly/monthly] highlights—[dd Month yyyy]'
                khminisummary = etree.SubElement(khbody, '{%s}mini-summary' % NSMAP['kh'])
                khminisummary.text = "This [week's/month's] edition of [PA] [weekly/monthly] highlights includes:"
                
                #print(topicList)
                for topic in topicList:
                    #print(topic)
                    dfTopic = dfPA[(dfPA.TopicName == topic)] 
                    topicCount = len(dfTopic)
                    dfTopic = dfTopic.sort_values(['Title'], ascending = True)
                    trsecmain = etree.SubElement(khbody, '{%s}secmain' % NSMAP['tr'])
                    coretitle = etree.SubElement(trsecmain, '{%s}title' % NSMAP['core'])
                    coretitle.text = str(topic)

                    for x in range(0,topicCount):                               
                        newsTitle = dfTopic.Title.iloc[x]     
                        newsCitation = dfTopic.Citation.iloc[x]
                        newsDate = dfTopic.IssueDate.iloc[x]
                        newsPubDate = dfTopic.PubDate.iloc[x]
                        newsMiniSummary = dfTopic.MiniSummary.iloc[x]
                        newsSources = dfTopic.Sources.iloc[x]
                        try: newsURLs = dfTopic.URLs.iloc[x]
                        except AttributeError: newsURLs = 'nan'
                        #News
                        trsecsub1 = etree.SubElement(trsecmain, '{%s}secsub1' % NSMAP['tr'])
                        coretitle = etree.SubElement(trsecsub1, '{%s}title' % NSMAP['core'])
                        coretitle.text = newsTitle
                        
                        #mini summary section:
                        paragraphs = re.split(r"\n\n", newsMiniSummary)
                        for para in paragraphs:
                            corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                            corepara.text = para

                        corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                        if 'urn:' in newsCitation: 
                            for item in ['Law360']: #, 'MLex'
                                if item in newsMiniSummary:
                                    corepara.text = 'See: '
                                else:
                                    corepara.text = 'See News Analysis: '
                            lncicite = etree.SubElement(corepara, '{%s}cite' % NSMAP['lnci'])
                            lncicite.set('normcite', newsCitation)
                            lncicontent = etree.SubElement(lncicite, '{%s}content' % NSMAP['lnci'])
                            lncicontent.text = newsTitle
                            lncicite.tail = '.'
                        else:  corepara.text = 'See: ' + newsCitation + '.'
                        print('newsURLs: ' + str(newsURLs))
                        if str(newsURLs) != 'nan':
                            if str(newsSources) != 'nan':
                                #newsSources = newsSources.strip('[]')
                                #newsSources = newsSources.replace('\\n', '') #convert string list to list
                                #newsSources = newsSources.replace("', '","'$?$ '" ) #convert string list to list
                                #newsSources = newsSources.strip("'")
                                #newsSources = newsSources.split("$?$")
                                #print(newsSources)
                                newsSources = etree.fromstring(newsSources)                       
                                newsSourcesLen = newsSources.xpath('count(//url)')
                                if newsSourcesLen > 1:                               
                                    corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                    corepara.text = 'Sources: ' 
                                    i=2
                                    newsSources = newsSources.findall(".//url")
                                    for newsSource in newsSources:
                                        coreurl = etree.SubElement(corepara, '{%s}url' % NSMAP['core'])
                                        coreurl.set('address', newsSource.get('address'))
                                        coreurl.text = newsSource.text
                                        #print(i, newsSourcesLen)
                                        if i < newsSourcesLen:
                                            coreurl.tail = ', '
                                        else:
                                            if i == newsSourcesLen:
                                                coreurl.tail = ' and '
                                            else: coreurl.tail = '.'
                                        i=i+1
                                else:
                                    try: 
                                        newsSource = newsSources.find(".//url")
                                        newsSourceAddress = str(newsSource.get('address'))
                                        corepara = etree.SubElement(trsecsub1, '{%s}para' % NSMAP['core'])
                                        corepara.text = 'Source: '                                            
                                        coreurl = etree.SubElement(corepara, '{%s}url' % NSMAP['core'])
                                        coreurl.set('address', newsSourceAddress)                                
                                        coreurl.text = newsSource.text
                                        coreurl.tail = '.'
                                    except AttributeError: print('No source URL given...')
                    
                            

                tree = etree.ElementTree(khdoc)
                tree.write(XMLFilepath,encoding='utf-8')

                f = open(XMLFilepath,'r', encoding='utf-8')
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
                f = open(XMLFilepath,'w', encoding='utf-8')
                f.write(newdata)
                f.close()

                print('XML exported to...' + XMLFilepath)
                LogOutput('XML exported to...' + XMLFilepath)

                #Insert news items into template file in the watched folder
                templateFilepath = FindMostRecentFile(outputDir + constantPA + '\\Auto-generate XML\\', '*.xml')            
                if templateFilepath != 'na': 
                    #templateFilepathDate = re.search('template (.*) test',templateFilepath).group(1) 
                    
                    postConversionTemplateFilename = re.search(r'.*\\(.*)', templateFilepath).group(1)
                    postConversionTemplateFilepath = outputDir + constantPA + '\\' + postConversionTemplateFilename

                    templatetree = etree.parse(templateFilepath).getroot()         
                    try:
                        templateKHbody = templatetree.find('.//kh:body', NSMAP)    #find element of elements where insertion will be   
                        i=0
                        for trsecmain in tree.findall('.//tr:secmain', NSMAP):
                            templateKHbody.insert(2+i, trsecmain) #insert at 3rd element after increment
                            i+=1
                    except:
                        print('Problem inserting news items into the template...')
                        LogOutput('Problem inserting news items into the template...')
                    tree = etree.ElementTree(templatetree)
                    tree.write(postConversionTemplateFilepath, encoding="utf-8", xml_declaration=True, doctype='<!DOCTYPE kh:document SYSTEM "\\\\voyager\\templates\\DTDs\\LNUK\\KnowHow\\KnowHow.dtd">')    
                    
                    print('Template XML converted and exported to...' + postConversionTemplateFilepath)
                    LogOutput('Template XML converted and exported to:\n' + postConversionTemplateFilepath)
                    os.remove(templateFilepath) #Delete old file
                    print('Deleted old template file: ' + templateFilepath)
                    LogOutput('Deleted old template file:\n' + templateFilepath)
                else:                    
                    print('No Template supplied for: ' + constantPA)
                    LogOutput('No Template supplied for: ' + constantPA)
                    
                if os.path.isdir(archiveDir) == False:
                    os.makedirs(archiveDir)

                shutil.copy(XLSFilepath, XLSarchiveFilepath) #Copy
                os.remove(XLSFilepath) #Delete old file
                print('Moved: ' + XLSFilepath + ', to: ' + XLSarchiveFilepath)
                LogOutput('Moved: ' + XLSFilepath + '\nto: ' + XLSarchiveFilepath)
                
        except AttributeError: 
            print('XLSX found but filename not in the expected format: ' + XLSFilepath)
            LogOutput('XLSX found but filename not in the expected format: ' + XLSFilepath)
    else:
        print('\nXSLX file not found in the watched folder for...' + constantPA)
        LogOutput('\nXSLX file not found in the watched folder for...' + constantPA)

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
        LogOutput('Problem loading template xml...')
        return 'section not found'

def FindMostRecentFile(directory, pattern):
    try:
        filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
        filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
        return filelist[0]
    except: return 'na'


def LogOutput(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()


#Directories
logDir = "\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Logs\\"   
#logDir = "\\\\atlas\\lexispsl\\Highlights\\dev\\Automatic creation\\Logs\\" 
outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'


AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    


JCSLogFile = logDir + 'JCSlog-newsxmlgen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()

highlightDate = str(time.strftime("%#d %B %Y"))
#main script
print("Today's date is: " + str(highlightDate))
LogOutput("Today's date is: " + str(highlightDate))

print("\nNews XML auto-generation for weekly highlights...\n")
LogOutput("\nNews XML auto-generation for weekly highlights...\n")
    
for PA in AllPAs:
    #if PA not in MonthlyPAs: 
    XMLGeneration(PA, highlightDate, 'weekly', outputDir)
        


print('Finished')
LogOutput('Finished')


