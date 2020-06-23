#AICER updated content log for highlights, loops through all the most recent AICER report and builds a log based on content updated/new in the last week/month
#last updated 20.05.20
#Developed by Daniel Hutchings
#JR added logfile info 19.08.19

import csv
import pandas as pd
import glob
import os
import fnmatch
import re
import time
import datetime
import sys
import shutil
import xml.etree.ElementTree as ET
from lxml import etree
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def LogOutput(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()

def Filter(reportDir, filename, df, dfshortcuts, highlightType, updatenewtype):    
    print(updatenewtype, highlightType)
    LogOutput(str(updatenewtype) + " " + str(highlightType))
    date =  str(time.strftime("%d%m%Y"))
    if highlightType == 'weekly': timeago = (datetime.datetime.now().date() - datetime.timedelta(8)) #the 'date' part of this means it will only provide the date, not the hours, min, sec etc
    if highlightType == 'monthly': timeago = (datetime.datetime.now().date() - datetime.timedelta(32))
    daterange = timeago.strftime('%b %d') + ' - ' + time.strftime('%b %d, %Y')

    
    df = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    dfshortcuts = dfshortcuts[dfshortcuts.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    
    df['LastPublishedDate'] = pd.to_datetime(df['LastPublishedDate'], dayfirst=False) #Date is in American format hence dayfirst false
    df['LastUnderReviewDate'] = pd.to_datetime(df['LastUnderReviewDate'], dayfirst=False) #Date is in American format hence dayfirst false

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

    if updatenewtype == 'updated':
        df1 = df
        df = df[df.MajorUpdateFirstPublished.notnull()]
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        
        
        print("Grabbing docs that were updated after: " + str(timeago))
        LogOutput("Grabbing docs that were updated after: " + str(timeago))
        df = df[df.MajorUpdateFirstPublished.dt.date > timeago]
        reportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_updated_" + date + ".csv"
   
    if updatenewtype == 'new':
        df1 = df
        df = df[df.DateFirstPublished.notnull()]
        df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist', 'QandAs'])]
        df = df[df.PageType.isin(['Precedents', 'Practical Guidance', 'Checklists', 'Q&As', 'InteractiveFlowchart', 'StaticFlowchart', 'SmartPrecedent'])] # These are the content types to keep in the report
        
        
        print("Grabbing docs that were created after: " + str(timeago))
        LogOutput("Grabbing docs that were created after: " + str(timeago))
        df = df[df.DateFirstPublished.dt.date > timeago]
        reportFilename = re.search('([^\.]*)\.csv',filename).group(1) + "_UKPSL_" + highlightType + "_HL_new_" + date + ".csv"
   

    #dropping unnecessary columns
    del df['DisplayId'], df['LexisSmartId'], df['OriginalContentItemPA'], df['TopicTreeLevel4'], df['TopicTreeLevel5'], df['TopicTreeLevel6'], df['TopicTreeLevel7'], df['TopicTreeLevel8'], df['TopicTreeLevel9'], df['TopicTreeLevel10'], df['VersionFilename'], df['Filename_Or_Address'], df['CreateDate'], df['LastPublishedDate'], df['OriginalLastPublishedDate'], df['LastMajorDate'], df['LastMinorDate'], df['LastReviewedDate'], df['LastUnderReviewDate'], df['SupportsMiniSummary'], df['HasMiniSummary']
    df = df.rename(columns={'id': 'DocID', 'TopicTreeLevel1': 'PA', 'TopicTreeLevel2': 'Subtopic', 'Label': 'DocTitle'})    
    df['Subtopic'] = df['Subtopic'] + ' > ' + df['TopicTreeLevel3'] 
    del df['TopicTreeLevel3']


    #searching for shortcuts
    print('Searching for shortcuts...')
    LogOutput("Searching for shortcuts...")
    i = 0
    #df1.to_csv(aicerDir + 'test-df1-' + updatenewtype + '-' + highlightType + '.csv', sep=',',index=False, encoding='utf-8')

    for index, row in df.iterrows():
        DocID = df.iloc[i,0]
        #print(DocID)
        DocTitle = df.DocTitle.iloc[i]
        masterPA = df.iloc[i,4]
        masterContentType = df.iloc[i,1]
        listofpas = [masterPA]
        
        #individual shortcuts
        if any(df1['OriginalContentItemId'] == DocID) == True:
            shortcutcount = len(df1[df1.OriginalContentItemId.isin([DocID])]) #filter by OCI with DocID to find how many shortcuts
            #print(shortcutcount)
            for x in range(0,shortcutcount):       
                PA = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel1'].iloc[x])
                #print(x, PA)
                if PA not in listofpas:    
                    
                    ContentItemType = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'ContentItemType'].iloc[x])
                    PageType = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'PageType'].iloc[x])
                    Subtopic = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel2'].iloc[x]) + ' > ' + str(df1.loc[df1['OriginalContentItemId'] == DocID, 'TopicTreeLevel3'].iloc[x])
                    #DocTitle = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'Label'].iloc[x])
                    
                    OriginalContentItemId = int(DocID)
                                   
                    DocID2 = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'id'].iloc[x])
                    #print('DocID2', DocID2)
                    UR = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'UnderReview'].iloc[x])
                    MajorUpdateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'MajorUpdateFirstPublished'].iloc[x])
                    DateFirstPublished = str(df1.loc[df1['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                    dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemType":masterContentType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA, "OriginalContentItemPA":masterPA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                    #print(dictionary_row)
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
                            #DocTitle = str(dfshortcuts.loc[dfshortcuts['id'] == DocID, 'Label'].iloc[x])
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
                            #DocTitle = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'Label'].iloc[x])
                            OriginalContentItemId = int(DocID)
                            DocID2 = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'id'].iloc[x])
                            UR = str(df1.loc[df1['id'] == DocID, 'UnderReview'].iloc[0])
                            MajorUpdateFirstPublished = 'nan'
                            DateFirstPublished = str(dfshortcuts.loc[dfshortcuts['OriginalContentItemId'] == DocID, 'DateFirstPublished'].iloc[x])
                            dictionary_row = {"DocID":DocID2,"ContentItemType":ContentItemType,"OriginalContentItemType":masterContentType,"OriginalContentItemId":OriginalContentItemId,"PageType":PageType,"PA":PA, "OriginalContentItemPA":masterPA,"Subtopic":Subtopic,"DocTitle":DocTitle,"DateFirstPublished":DateFirstPublished,"MajorUpdateFirstPublished":MajorUpdateFirstPublished,"UnderReview":UR}
                            df = df.append(dictionary_row, ignore_index=True)                    
                            listofpas.append(PA)
                                                
        i=i+1

    print('Total found: ' + str(len(df)))
    LogOutput("Total found: " + str(len(df)))
    #cleanup
    df.OriginalContentItemId = df.OriginalContentItemId.fillna(0) #fill all empty values of this column with zeros
    df.OriginalContentItemId = df.OriginalContentItemId.astype(int) #convert hidden float column to int to remove trailing decimals when exporting to csv

    if updatenewtype == 'updated':
        columnsTitles = ['DocID', 'OriginalContentItemId', 'ContentItemType', 'OriginalContentItemType', 'PageType', 'PA', 'OriginalContentItemPA', 'Subtopic', 'DocTitle', 'MajorUpdateFirstPublished', 'UnderReview']
    
    if updatenewtype == 'new':
        columnsTitles = ['DocID', 'OriginalContentItemId', 'ContentItemType', 'OriginalContentItemType', 'PageType', 'PA', 'OriginalContentItemPA', 'Subtopic', 'DocTitle', 'DateFirstPublished', 'UnderReview']
    
    #reorder columns by list of columnTitles
    df = df.reindex(columns=columnsTitles)

    df.to_csv(reportDir + reportFilename, sep=',',index=False, encoding='utf-8')
    print('Exported to ' + reportDir + reportFilename)
    LogOutput("Exported to " + str(reportDir) + " " + str(reportFilename))
    return(reportDir + reportFilename)



def FindMostRecentFile(directory, pattern):
    filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
    filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
    return filelist[0]


def FindRecentFile(directory, pattern, number):
    filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
    filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
    return filelist[number]


def ListOfFilesInDirectory(directory, pattern):
    TodaysDate = datetime.date.today()
    existingList = []
    filelist = glob.iglob(os.path.join(directory, pattern)) #builds list of file in a directory based on a pattern
    for filepath in filelist:
        if TodaysDate != datetime.datetime.fromtimestamp(os.path.getmtime(filepath)).date(): 
            existingList.append(filepath)
        else:
            print('Following file has date the same as today so not putting it on the To Archive list: ' + filepath)

    return existingList

def Archive(listOfFiles, reportDir):
    copyList = []
    archiveDir = reportDir + 'Archive\\'

    #Check archive folder exists
    if os.path.isdir(archiveDir) == False:
        os.makedirs(archiveDir)

    for existingFilepath in listOfFiles:
        print(existingFilepath)
        directory = re.search('(.*\\\\)[^\.]*\.csv',existingFilepath).group(1)
        filename = re.search('.*\\\\([^\.]*\.csv)',existingFilepath).group(1)
        destinationFilepath = directory + 'Archive\\' + filename
        shutil.copy(existingFilepath, destinationFilepath) #Copy
        os.remove(existingFilepath) #Delete old file
        copyList.append('Moved: ' + existingFilepath + ', to: ' + destinationFilepath)
    return copyList



def formatEmail(reciever_email, subject, html):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email

    HTMLPart = MIMEText(html, "html")
    msg.attach(HTMLPart)
    return msg

def sendEmail(msg, receiver_email):
    s = smtplib.SMTP("LNGWOKEXCP002.legal.regn.net")
    s.sendmail(sender_email, receiver_email, msg.as_string())


#main script
#reportDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Reports\\'
reportDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\New and Updated content report\\'
#reportDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\'
#reportDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\"
logDir = "\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Logs\\"
#logDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\Highlights\\Logs\\'
aicerDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER\\'
#aicerDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\"
globalmetricsDir = '\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER_withShortcut_AdHoc\\'
shortcutNodeDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER_Shortcuts\\'
#globalmetricsDir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER_Shortcuts\\'
#globalmetricsDir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\"
pguidlistDir = '\\\\lngoxfdatp16vb\\Fabrication\\MasterStore\\PGUID-Lists\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\lookup-dpsis.csv'


existingReports = ListOfFilesInDirectory(reportDir, '*.csv')

JCSLogFile = logDir + 'JCSlog-reportgen.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()

print("New and Updated content report for highlights...")
LogOutput("New and Updated content report for highlights...")

aicerMostRecent = FindRecentFile(aicerDir, '*AICER*.csv', 1)
aicerSecondMostRecent = FindRecentFile(aicerDir, '*AICER*.csv', 2) #1 is the second on the list

aicerMostRecentSize = os.path.getsize(aicerMostRecent)
aicerSecondMostRecentSize = os.path.getsize(aicerSecondMostRecent)
print(aicerMostRecent, aicerMostRecentSize)
print(str(aicerMostRecentSize + 2000000))
print(aicerSecondMostRecent, aicerSecondMostRecentSize)

if aicerMostRecentSize + 2000000 < aicerSecondMostRecentSize: #add 2000kb as it's ok to have the aicer slightly less, but not less than a couple of meg
    print("Problem with the most recent Aicer report: filesize appears to be significantly less than the previous Aicer report which indicates an error. Manual investigation required...")
    errorTitle = "Problem with the most recent Aicer report: filesize appears to be significantly less than the previous Aicer report which indicates an error. Manual investigation required...\n"
    text1 = "Most recent Aicer: " + aicerMostRecent + " Size: " + str(aicerMostRecentSize) + "\n"
    text2 = "Second most recent Aicer: " + aicerSecondMostRecent + " Size: " + str(aicerSecondMostRecentSize)    
    LogOutput(errorTitle + text1 + text2)

    html = "<html><head><title>" + errorTitle + "</title></head><body>"
    html += '<div style="font-size: 100%; text-align: center;">'
    html += '<p>' + errorTitle + '</p>'
    html += "<p><b>Most recent Aicer: </b><br />" + aicerMostRecent + " <br /><b>Size: </b>" + str(aicerMostRecentSize) + "</p>"
    html += "<p><b>Second most recent Aicer: </b><br />" + aicerSecondMostRecent + " <br /><b>Size: </b>" + str(aicerSecondMostRecentSize) + "</p>"
    html += 'Investigate further in the below locations: <br />'
    html += '<p><a href="file://///atlas/lexispsl/Highlights/Automatic creation/AICER/" target="_blank">\\\\atlas\lexispsl\Highlights\Automatic creation\AICER\</a></p>'
    html += '<p><a href="file://///atlas/knowhow/PSL_Content_Management/AICER_Reports/AICER_PM/" target="_blank">\\\\atlas\knowhow\PSL_Content_Management\AICER_Reports\AICER_PM\</a></p>'
    html += '<sup>See the log for more details: <br />'
    html += '<a href="file://///atlas/lexispsl/Highlights/Automatic creation/Logs/JCSlog-reportgen.txt" target="_blank">\\\\atlas\lexispsl\Highlights\Automatic creation\Logs\JCSlog-reportgen.txt</a></sup>'
    html += '</div></body></html>'

    #Email section
    sender_email = 'LNGUKPSLDigitalEditors@ReedElsevier.com'
    receiver_email_list = ['daniel.hutchings.1@lexisnexis.co.uk', 'james-john.dwyer-wilkinson@lexisnexis.co.uk', 'LNGUKPSLDigitalEditors@ReedElsevier.com']

    #create and send email
    for receiver_email in receiver_email_list:
        msg = formatEmail(receiver_email, errorTitle, html)
        sendEmail(msg, receiver_email)
    print('Email sent...')
    LogOutput('Email sent to...' + str(receiver_email_list))

else:
    aicerFilename = re.search('.*\\\\AICER\\\\([^\.]*\.csv)',aicerMostRecent).group(1)
    print('Loading the most recent AICER report: ' + aicerFilename)
    LogOutput("Loading the most recent AICER report: " + str(aicerFilename))
    aicershortcutsFilename = FindMostRecentFile(shortcutNodeDir, 'AllContentItemsExportWithShortCutNodeInfo*.csv')
    print('Loading the most recent AICER Shortcuts report: ' + aicershortcutsFilename)
    LogOutput("Loading the most recent AICER Shortcuts report: " + str(aicershortcutsFilename))

    #filter
    dfaicer = pd.read_csv(aicerDir + aicerFilename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
    print('Aicer loaded...loading Aicer shortcuts...')
    LogOutput("Aicer loaded...loading Aicer shortcuts...")
    dfshortcuts =  pd.read_csv(aicershortcutsFilename, encoding='utf-8', low_memory=False) #Load csv file into dataframe
    print('Aicer shortcuts loaded...filtering reports...')
    LogOutput("Aicer shortcuts loaded...filtering reports...")

    Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'weekly', 'new')
    Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'weekly', 'updated')
    Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'monthly', 'new')
    Filter(reportDir, aicerFilename, dfaicer, dfshortcuts, 'monthly', 'updated')

    archivedList = Archive(existingReports, reportDir)
    print(archivedList)
    LogOutput('Moved following files to Archive folder: ' + str(archivedList))


LogOutput("End")
