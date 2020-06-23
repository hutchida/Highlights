import zipfile
import re
import datetime
import xml.etree.ElementTree as ET
from lxml import etree
import glob
import os
import time
import datetime

def GrabTemplatesFromPSLDump(zipFilepath, lastWeekDate, lastWeeksBrexitDate):
    zfile = zipfile.ZipFile(zipFilepath)    
    k=0
    for finfo in zfile.infolist():
        if '0S4D' in finfo.filename:
            ifile = zfile.open(finfo) 
            line_list = ifile.read()
            tree = etree.fromstring(line_list, etree.XMLParser(remove_pis=True))
        
            khdoctitle = tree.find('.//kh:document-title', NSMAP)
            if khdoctitle is not None:
                if khdoctitle.text is not None:
                    for PA in AllPAs:
                        exportFilepath =  outputDir + PA + '\\' + PA + ' weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'Life Sciences and Pharmaceuticals': exportFilepath =  outputDir + 'Life Sciences\\Life Sciences weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'Banking &amp; Finance': exportFilepath =  outputDir + 'Banking and Finance\\Banking and Finance weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'Banking & Finance': exportFilepath =  outputDir + 'Banking and Finance\\Banking and Finance weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'In-house': exportFilepath =  outputDir + 'In-House Advisor\\In-House Advisor weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'In-House': exportFilepath =  outputDir + 'In-House Advisor\\In-House Advisor weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'Share Incentives': exportFilepath =  outputDir + 'Share Schemes\\Share Schemes weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'Life Sciences': exportFilepath =  outputDir + 'Life Sciences and Pharmaceuticals\\Life Sciences weekly highlights ' + lastWeekDate + '_preview.xml'
                        if PA == 'Life Sciences and Pharmaceuticals': exportFilepath =  outputDir + 'Life Sciences and Pharmaceuticals\\Life Sciences weekly highlights ' + lastWeekDate + '_preview.xml'
                        
                        if PA == 'Brexit':
                            exportFilepath =  outputDir + PA + '\\' + PA + ' highlights ' + lastWeeksBrexitDate + '_preview.xml'                        
                            if PA + ' highlights—' + lastWeeksBrexitDate in khdoctitle.text:
                                khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                                khdoc.append(tree.find('.//kh:body', NSMAP))
                                tree = etree.ElementTree(khdoc)
                                try: tree.write(exportFilepath,encoding='utf-8')
                                except FileNotFoundError: 
                                    print('FileNotFound: ' + exportFilepath)
                                    LogOutput('FileNotFound: ' + exportFilepath)
                                LogOutput('Exported: ' + exportFilepath)
                                print('Exported: ' + exportFilepath)
                                k+=1
                            else: 
                                if PA + ' highlights&#x2014;' + lastWeeksBrexitDate in khdoctitle.text:
                                    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                                    khdoc.append(tree.find('.//kh:body', NSMAP))
                                    tree = etree.ElementTree(khdoc)
                                    tree.write(exportFilepath,encoding='utf-8')
                                    LogOutput('Exported: ' + exportFilepath)
                                    print('Exported: ' + exportFilepath)
                                    k+=1
                        else:
                            if PA + ' weekly highlights—' + lastWeekDate in khdoctitle.text:
                                khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                                khdoc.append(tree.find('.//kh:body', NSMAP))
                                tree = etree.ElementTree(khdoc)
                                try: tree.write(exportFilepath,encoding='utf-8')
                                except FileNotFoundError: 
                                    print('FileNotFound: ' + exportFilepath)
                                    LogOutput('FileNotFound: ' + exportFilepath)
                                LogOutput('Exported: ' + exportFilepath)
                                print('Exported: ' + exportFilepath)
                                k+=1
                            else: 
                                if PA + ' weekly highlights&#x2014;' + lastWeekDate in khdoctitle.text:
                                    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                                    khdoc.append(tree.find('.//kh:body', NSMAP))
                                    tree = etree.ElementTree(khdoc)
                                    tree.write(exportFilepath,encoding='utf-8')
                                    LogOutput('Exported: ' + exportFilepath)
                                    print('Exported: ' + exportFilepath)
                                    k+=1
                                
                
    print(k)
    LogOutput(str(k) + ' highlights docs from last week found')
        

def FindMostRecentFile(directory, pattern):
    try:
        filelist = os.path.join(directory, pattern) #builds list of file in a directory based on a pattern
        filelist = sorted(glob.iglob(filelist), key=os.path.getmtime, reverse=True) #sort by date modified with the most recent at the top
        return filelist[0]
    except: return 'na'            


def FindLastWeekday(givendate, weekday):
    givendate += datetime.timedelta(days=-1) #incase this is run on a thursday, just pretend the day before so you get last thursday instead of today which is a thursday, hence minus 1
    dayshift = (givendate.weekday() - weekday) % 7 #find difference in days between day of the week given and the supplied date
    return givendate - datetime.timedelta(days=dayshift) #subtract that difference from the date given to give you last thursday, or whatever day supplied


def LogOutput(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()


AllPAs = ['Arbitration', 'Banking &amp; Finance', 'Banking and Finance', 'Banking & Finance', 'Brexit', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'In-house', 'In-House', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Life Sciences', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Share Incentives', 'Tax', 'TMT', 'Wills & Probate', 'Wills &amp; Probate', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
NSMAP = {'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'em': 'http://www.lexisnexis.com/namespace/sslrp/em', 'fi': 'http://www.lexisnexis.com/namespace/sslrp/fi', 'fm': 'http://www.lexisnexis.com/namespace/sslrp/fm', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'form': 'http://www.lexisnexis.com/namespace/sslrp/form', 'glph': 'http://www.lexisnexis.com/namespace/sslrp/glph', 'header': 'http://www.lexisnexis.com/namespace/sslrp/header', 'ic': 'http://www.lexisnexis.com/namespace/sslrp/ic', 'in': 'http://www.lexisnexis.com/namespace/sslrp/in', 'lnb-case': 'http://www.lexisnexis.com/namespace/case/lnb-case', 'lnb-leg': 'http://www.lexisnexis.com/namespace/sslrp/lnb-leg', 'lnb-nl': 'http://www.lexisnexis.com/namespace/sslrp/lnb-nl', 'lnb-pg': 'http://www.lexisnexis.com/namespace/uk/lnb-practicalguidance', 'lnbdig-case': 'http://www.lexisnexis.com/namespace/digest/lnbdig-case', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'm': 'http://www.w3.org/1998/Math/MathML', 'se': 'http://www.lexisnexis.com/namespace/sslrp/se', 'su': 'http://www.lexisnexis.com/namespace/sslrp/su', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance', 'atict': 'http://www.arbortext.com/namespace/atict'}
#outputDir = '\\\\atlas\\lexispsl\\Highlights\\Practice Areas\\'
outputDir = '\\\\atlas\\lexispsl\\Highlights\\dev\\Practice Areas\\'
logDir = "\\\\atlas\\lexispsl\\Highlights\\Automatic creation\\Logs\\"
zipDir = "\\\\voyager\\edit_systems\\Content Management Team\\"


JCSLogFile = logDir + 'JCSlog-QueryPSLZip.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y"))
l.write("Start "+logdate+"\n")
l.close()

zipFilepath = FindMostRecentFile(zipDir, '20*.zip')
givendate = datetime.datetime.today()
#givendate = datetime.date(2019, 12, 26)
lastThursday = FindLastWeekday(givendate, 3) # 3 is thursday
lastFriday = FindLastWeekday(givendate, 4)
lastWeeksDate = str(lastThursday.strftime("%#d %B %Y"))
lastWeeksBrexitDate = str(lastFriday.strftime("%#d %B %Y"))

print('Searching for highlights with the following date in the title: ' + lastWeeksDate + ' ... and the following date for Brexit: ' + lastWeeksBrexitDate)
LogOutput('Searching for highlights with the following date in the title: ' + lastWeeksDate + ' ... and the following date for Brexit: ' + lastWeeksBrexitDate)
print('zipFilepath: ' + zipFilepath)
LogOutput('zipFilepath: ' + zipFilepath)
start = datetime.datetime.now() 
GrabTemplatesFromPSLDump(zipFilepath, lastWeeksDate, lastWeeksBrexitDate)
print("Finished! That took..." + str(datetime.datetime.now() - start))
LogOutput("Finished! That took..." + str(datetime.datetime.now() - start))