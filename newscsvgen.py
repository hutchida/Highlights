#News CSV generator for highlights, loops through all entries on the news items list supplied by Martin/Connell and creates a CSV file for each PA in their respective folder. PSLs will then add info to these spreadsheets so that the newsxmlgen script can be instructed correctly
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

    
def CSVGeneration(PA, highlightDate, highlightType, df, outputDir):
    constantPA = PA            
    CSVFilepath = outputDir + constantPA + '\\' + constantPA + ' news items ' + highlightDate + '.csv'
    dfPA = df[(df.PA ==PA)] 
    dfPA = dfPA.sort_values(['Title'], ascending = True)
    dfPA.to_csv(CSVFilepath, index=False)
    print('CSV exported to...' + CSVFilepath)
    LogOutput('CSV exported to...' + CSVFilepath)
   
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
reportDir = '\\\\atlas\\Knowhow\\AutomatedContentReports\\NewsReport\\'
#reportDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\'

logDir = "\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Logs\\"    
#logDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\Logs\\'
#outputDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\xml\\Practice Areas\\'
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    

JCSLogFile = logDir + 'JCSlog-newscsvgen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()

highlightDate = str(time.strftime("%#d %B %Y")) #the hash character turns off the leading zero in the day   

#main script
print("Today's date is: " + highlightDate)
LogOutput("Today's date is: " + highlightDate)

newsListFilepath = FindMostRecentFile(reportDir, '*.xml')

df = pd.DataFrame()
LogOutput('Looping through the news xml log...')
newsListXML = etree.parse(newsListFilepath)
newsItems = newsListXML.findall(".//document")
for newsItem in newsItems:
    newsSourceList = []
    newsTitle = newsItem.find(".//title")
    newsTitle = newsTitle.text
    newsCitation = newsItem.find(".//citation")
    newsCitation = newsCitation.text
    newsMiniSummary = newsItem.find(".//mini-summary")
    newsMiniSummary = newsMiniSummary.text
    newsMiniSummary = re.sub("^.*analysis: ", "", newsMiniSummary)    
    newsMiniSummary = re.sub("^.*Analysis: ", "", newsMiniSummary)    
    newsDate = newsItem.find(".//date")
    newsDate = newsDate.text
    try: newsSources = etree.tostring(newsItem.find(".//source-links"), encoding="unicode")
    except TypeError: newsSources = ''
    newsPAs = newsItem.findall(".//practice-area")
    for newsPA in newsPAs:
        dictionary_row = {"Title":newsTitle,"Citation":newsCitation,"MiniSummary":newsMiniSummary,"Date":newsDate,"Sources":newsSources,"PA":newsPA.text}
        df = df.append(dictionary_row, ignore_index=True)        
        LogOutput(str(newsCitation))    
        print(str(newsCitation))
df['TopicName'] = ''
columnsTitles = ['TopicName', 'Title', 'MiniSummary', 'Date', 'Sources', 'Citation', 'PA']
df = df.reindex(columns=columnsTitles) #reorder columns
df.to_csv(logDir + 'all-pas-news-list.csv', index=False)


LogOutput("\nNews CSV guide generation for weekly highlights...\n")
 
for PA in AllPAs:
    if PA not in MonthlyPAs: 
        CSVGeneration(PA, highlightDate, 'weekly', df, outputDir)

print('Finished, access the log here: ' + logDir + 'JCSlog-newscsvgen.txt')
LogOutput('Finished')


