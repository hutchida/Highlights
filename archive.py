#Archive script, to be run every day, depending on the day will archive different locations
#Developed by Daniel Hutchings

import glob
import os
import re
import time
import datetime
import calendar
import sys
import shutil


def log(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    print(message)
    l.close()

def DayOfTheWeek(givendate):
    if givendate.weekday() == 0: return 'Monday'
    if givendate.weekday() == 1: return 'Tuesday'
    if givendate.weekday() == 2: return 'Wednesday'
    if givendate.weekday() == 3: return 'Thursday'
    if givendate.weekday() == 4: return 'Friday'
    if givendate.weekday() == 5: return 'Saturday'
    if givendate.weekday() == 6: return 'Sunday'

def ListOfFilesInDirectory(directory, pattern):
    existingList = []
    filelist = glob.iglob(os.path.join(directory, pattern)) #builds list of file in a directory based on a pattern
    for filepath in filelist: existingList.append(filepath)        
    return existingList


def Archive(listOfFiles, archiveDir):
    copyList = []
    
    #Check archive folder exists
    if os.path.isdir(archiveDir) == False:
        os.makedirs(archiveDir)

    for existingFilepath in listOfFiles:
        print(existingFilepath)
        directory = re.search('(.*\\\\)[^\.]*\..*',existingFilepath).group(1)
        filename = re.search('.*\\\\([^\.]*\..*)',existingFilepath).group(1)
        destinationFilepath = directory + 'Archive\\' + filename
        shutil.copy(existingFilepath, destinationFilepath) #Copy
        try: 
            os.remove(existingFilepath) #Delete old file        
            copyList.append('Moved: ' + existingFilepath + ', to: ' + destinationFilepath)
        except:
            log('Could not delete: ' + existingFilepath)
    return copyList


def find_most_recent_file(directory, pattern):
    try:
        filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
        filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
        return filelist[0]
    except: return 'Error finding most recent file'

def backup_aicer(aicer_dir, aicer_backup_dir, pattern):   
    log('Backing up: ' + pattern)
    current_aicer_filepath = aicer_dir + pattern + '.csv'
    backup_aicer_filepath = aicer_backup_dir + pattern + '-' + str(time.strftime("%Y%m%d")) + '.csv'
    #print(current_aicer_filepath)
    #print(backup_aicer_filepath)
    try: 
        shutil.copy(current_aicer_filepath, backup_aicer_filepath) #Copy
        log('Created dated version of most recent csv file in this directory: ' + backup_aicer_filepath)
    except: log('Error backing up current aicer...')
    #removing oldest csv file in the directory
    backup_list = os.path.join(aicer_backup_dir, '*.csv')
    backup_list = sorted(glob.iglob(backup_list), key=os.path.getmtime, reverse=False) #sort by date modified with the least recent at the top
    #log(str(backup_list))
    log(str(len(backup_list)))
    if len(backup_list) > 7:
        #log(str(backup_list))
        oldest_aicer = backup_list[0]
        log('Deleting the oldest csv in the backup: ' + oldest_aicer)
        try: os.remove(oldest_aicer) #Delete old file        
        except: log('Could not delete: ' + oldest_aicer)

state = 'live'
#state = 'livedev'

aicer_dir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER\\'
aicer_shortcuts_dir = '\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\AICER_Shortcuts\\'


if state == 'live':
    logDir = "\\\\atlas\\lexispsl\\Highlights\\logs\\"
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
    givendate = datetime.datetime.today()
    givenstrdate =  str(time.strftime("%Y-%m-%d"))    
    highlightsArchiveDir = '\\\\atlas\\Knowhow\\HighlightsArchive\\'
if state == 'livedev': 
    logDir = "\\\\atlas\\lexispsl\\Highlights\\dev\\logs\\"
    outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'
    #givendate = datetime.date(2020, 4, 2)
    #givenstrdate = "2020-03-27"
    givendate = datetime.datetime.today()
    givenstrdate =  str(time.strftime("%Y-%m-%d"))    
    highlightsArchiveDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\HighlightsArchive\\'

AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
exceptionList = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax'] 

JCSLogFile = logDir + 'JCSlog-archive.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d/%m/%Y %H:%M:%S"))
l.write("Start "+logdate+"\n")
l.close()

#THESE TASKS WILL BE DONE EVERY TIME THE SCRIPT IS RUN

#Backup aicer files for highlights project
backup_aicer(aicer_dir, aicer_dir, 'AICER')
backup_aicer(aicer_shortcuts_dir, aicer_shortcuts_dir, 'AllContentItemsExportWithShortCutNodeInfo')

#Archive anything found in the Law360 output folder
filesToArchive = ListOfFilesInDirectory("\\\\atlas\\lexispsl\\Automation\\Law360-conversion\\Output\\", '*.xml')
if len(filesToArchive) > 1:
    archivedList = Archive(filesToArchive, "\\\\atlas\\lexispsl\\Automation\\Law360-conversion\\Output\\archive\\")
    log('\nArchived:\n' + str(archivedList))
else:
    log('Nothing to archive in Law360 output folder...\n')


#Archive anything found in the Law360 DEV output folder
filesToArchive = ListOfFilesInDirectory("\\\\atlas\\lexispsl\\Automation-DEV\\Law360-conversion\\Output\\", '*.xml')
if len(filesToArchive) > 1:
    archivedList = Archive(filesToArchive, "\\\\atlas\\lexispsl\\Automation-DEV\\Law360-conversion\\Output\\archive\\")
    log('\nArchived:\n' + str(archivedList))
else:
    log('Nothing to archive in Law360 output folder...\n')



#THESE WILL BE DONE ONLY ON THE DAYS TESTED FOR

if DayOfTheWeek(givendate) == 'Friday':     
    print('Today is Friday...')
    log('Today is Friday...')
    #archive everything in the highlights PA folders that isn't a template file
    print('Archiving everything except template files in all PA highlights folders...')
    log('Archiving everything except template files in all PA highlights folders...')
    for PA in AllPAs:    
        if PA not in exceptionList:
            print(PA)
            log(PA)
            paDir = outputDir + PA
            archiveDir = outputDir + PA + '\\archive\\'
            filesToArchive = ListOfFilesInDirectory(paDir, '*.*') #get everything in the diretory
            filteredList = []
            for file in filesToArchive:
                if 'template' not in file: filteredList.append(file)
            filesToArchive = filteredList
            if len(filesToArchive) > 1:
                archivedList = Archive(filesToArchive, archiveDir)
                print('\nArchived:\n' + str(archivedList))
                log('\nArchived:\n' + str(archivedList))
            else:
                print('Nothing to archive...\n')
                log('Nothing to archive...\n')

            #wait = input("PAUSED...when ready press enter")


if DayOfTheWeek(givendate) == 'Thursday':     
    print('Today is Thursday...')
    log('Today is Thursday...')    
    print('Archiving highlights archive before repopulation later today...')
    log('Archiving highlights archive before repopulation later today...')

    archiveDir = highlightsArchiveDir + '\\archive\\'
    filesToArchive = ListOfFilesInDirectory(highlightsArchiveDir, '*.*') #get everything in the diretory
    if len(filesToArchive) > 1:
        archivedList = Archive(filesToArchive, archiveDir)
        print('\nArchived:\n' + str(archivedList))
        log('\nArchived:\n' + str(archivedList))
    else:
        print('Nothing to archive...\n')
        log('Nothing to archive...\n')
