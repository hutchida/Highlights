import zipfile
import re
import datetime
import xml.etree.ElementTree as ET
from lxml import etree
import glob
import os
import time
import datetime

import pandas as pd

def extract_info_from_zip(zip_filepath):       
    log('Extracting info from zip...')
    df = pd.DataFrame()        
    df_links = pd.DataFrame()    
    list_of_links=[]
    list_of_refs=[]
    zfile = zipfile.ZipFile(zip_filepath)
    i = 0
    for finfo in zfile.infolist():
        ifile = zfile.open(finfo)                 
        line_list = ifile.read()
        #GET DPSI, CHECK IT'S ON THE DPSI LIST IN ORDER TO PROCEED
        try: dpsi = re.search('([^\/]*)\/([^\.]*\.xml)',str(ifile.name)).group(1)
        except: dpsi = ''
        if dpsi not in dpsi_list:
            #print('Skipping...')
            continue #skip this loop if doc not on valid dpsi list
        
        try: doc_id = re.search('_([\d]*)\.xml', str(ifile.name)).group(1)
        except: doc_id = finfo

        try: tree = etree.fromstring(line_list, etree.XMLParser(remove_pis=True))
        except: print('Error accessing doc...')
                
        links = tree.xpath('//lnci:cite[contains(@normcite, "_")]', namespaces=NSMAP)
        if len(links)>0:
            for link in links:
                normcite = link.get('normcite')
                #print(normcite)
                links_row = {"normcite":normcite, "doc_id":doc_id}
                #log(str(links_row))
                #df_links = df_links.append(links_row, ignore_index=True) 
                list_of_links.append(links_row)
        #references = tree.findall('.//lnb-prec:contentReference[@name="id"]', NSMAP)
        references = tree.findall('.//lnb-prec:contentReference', NSMAP)
        if len(references)>0:
            for reference in references:
                dictionary_row = {"clause_id":reference.get('id'), "doc_id":doc_id}
                #log(str(dictionary_row))
                #df = df.append(dictionary_row, ignore_index=True) 
                list_of_refs.append(dictionary_row)
        df = pd.DataFrame(list_of_refs)
        df_links = pd.DataFrame(list_of_links)
    return df, df_links
            

            

def GrabTemplatesFromPSLDump(zip_filepath, lastWeekDate, lastWeeksBrexitDate):
    zfile = zipfile.ZipFile(zip_filepath)    

    
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
                                    log('FileNotFound: ' + exportFilepath)
                                log('Exported: ' + exportFilepath)
                                print('Exported: ' + exportFilepath)
                                k+=1
                            else: 
                                if PA + ' highlights&#x2014;' + lastWeeksBrexitDate in khdoctitle.text:
                                    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                                    khdoc.append(tree.find('.//kh:body', NSMAP))
                                    tree = etree.ElementTree(khdoc)
                                    tree.write(exportFilepath,encoding='utf-8')
                                    log('Exported: ' + exportFilepath)
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
                                    log('FileNotFound: ' + exportFilepath)
                                log('Exported: ' + exportFilepath)
                                print('Exported: ' + exportFilepath)
                                k+=1
                            else: 
                                if PA + ' weekly highlights&#x2014;' + lastWeekDate in khdoctitle.text:
                                    khdoc = etree.Element('{%s}document' % NSMAP['kh'], nsmap=NSMAP)
                                    khdoc.append(tree.find('.//kh:body', NSMAP))
                                    tree = etree.ElementTree(khdoc)
                                    tree.write(exportFilepath,encoding='utf-8')
                                    log('Exported: ' + exportFilepath)
                                    print('Exported: ' + exportFilepath)
                                    k+=1
                                
                
    print(k)
    log(str(k) + ' highlights docs from last week found')
        

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


def log(message):
    l = open(JCSLogFile,'a')
    l.write(message + '\n')
    l.close()
    print(message)


AllPAs = ['Arbitration', 'Banking &amp; Finance', 'Banking and Finance', 'Banking & Finance', 'Brexit', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'In-house', 'In-House', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Life Sciences', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Share Incentives', 'Tax', 'TMT', 'Wills & Probate', 'Wills &amp; Probate', 'Wills and Probate']    
MonthlyPAs = ['Competition', 'Family', 'Immigration', 'Insurance', 'Practice Compliance', 'Restructuring and Insolvency', 'Risk and Compliance']    
NSMAP = {'kh': 'http://www.lexisnexis.com/namespace/uk/kh', 'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'em': 'http://www.lexisnexis.com/namespace/sslrp/em', 'fi': 'http://www.lexisnexis.com/namespace/sslrp/fi', 'fm': 'http://www.lexisnexis.com/namespace/sslrp/fm', 'fn': 'http://www.lexisnexis.com/namespace/sslrp/fn', 'form': 'http://www.lexisnexis.com/namespace/sslrp/form', 'glph': 'http://www.lexisnexis.com/namespace/sslrp/glph', 'header': 'http://www.lexisnexis.com/namespace/sslrp/header', 'ic': 'http://www.lexisnexis.com/namespace/sslrp/ic', 'in': 'http://www.lexisnexis.com/namespace/sslrp/in', 'lnb-case': 'http://www.lexisnexis.com/namespace/case/lnb-case', 'lnb-leg': 'http://www.lexisnexis.com/namespace/sslrp/lnb-leg', 'lnb-nl': 'http://www.lexisnexis.com/namespace/sslrp/lnb-nl', 'lnb-pg': 'http://www.lexisnexis.com/namespace/uk/lnb-practicalguidance', 'lnbdig-case': 'http://www.lexisnexis.com/namespace/digest/lnbdig-case', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'm': 'http://www.w3.org/1998/Math/MathML', 'se': 'http://www.lexisnexis.com/namespace/sslrp/se', 'su': 'http://www.lexisnexis.com/namespace/sslrp/su', 'tr': 'http://www.lexisnexis.com/namespace/sslrp/tr', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance', 'atict': 'http://www.arbortext.com/namespace/atict'}
dpsi_list = ["0R2V","0R2W","0QOD","0OJ7","0OJ8","0OKE","0OJD","0OJE","0OKF","0OJJ","0OJK","0OJP","0OJQ","0OK5","0OK7","0OJV","0OJW","0OKK","0GVV","0OL7","0OLB","0OKN","0OLH","0OLI","0OKO","0OLM","0OLN","0OM7","0OMA","0OKQ","0OMH","0OMI","0OKR","0OMR","0OMS","0OKS","0ONR","0ONS","0ONI","0ONJ","0OKV","0OND","0ONE","0OKW","0ON8","0ON9","0OKX","0ON3","0ON4","0OKY","0OMW","0OMX","0OKZ","0R49","0R4A","0OMM","0OM8","0OM9","0OL6","0SBW","0SBX","0SBY","0OM2","0OM3","0OL8","0RCD","0RCE","0RCF","0OLU","0OLV","0OLP","0OLQ","0OLA","0X01","0X02","0XWH","0XWI"]

log_dir = r'\\atlas\Knowhow\Teleporting\logs'
zip_dir = "\\\\voyager\\edit_systems\\Content Management Team\\"

output_dir = r'\\atlas\Knowhow\Teleporting\reports'
links_dir = r'\\atlas\Knowhow\LinkHub'
JCSLogFile = log_dir + r'\JCSlog-QueryPSLZip.txt'
l = open(JCSLogFile,'w')
logdate =  str(time.strftime("%d%m%Y %H:%M:%S"))
l.write("Start "+logdate+"\n")
l.close()

zip_filepath = FindMostRecentFile(zip_dir, '20*.zip')
#zip_filepath = r'C:\Users\Hutchida\Documents\PSL\2020-05-09_08-00-47.zip'

givendate = datetime.datetime.today()
#givendate = datetime.date(2019, 12, 26)
lastThursday = FindLastWeekday(givendate, 3) # 3 is thursday
lastFriday = FindLastWeekday(givendate, 4)
lastWeeksDate = str(lastThursday.strftime("%#d %B %Y"))
lastWeeksBrexitDate = str(lastFriday.strftime("%#d %B %Y"))
NSMAP = {'xi': 'http://www.w3.org/2001/XInclude', 'lnci': 'http://www.lexisnexis.com/namespace/common/lnci', 'lnb-prec-cp': 'http://www.lexisnexis.com/namespace/uk/precedent-cp', 'header': 'http://www.lexisnexis.com/namespace/uk/header', 'lnb-prec': 'http://www.lexisnexis.com/namespace/uk/precedent', 'core': 'http://www.lexisnexis.com/namespace/sslrp/core', 'atict': 'http://www.arbortext.com/namespace/atict', 'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

log('zip_filepath: ' + zip_filepath)
start = datetime.datetime.now() 
#GrabTemplatesFromPSLDump(zip_filepath, lastWeeksDate, lastWeeksBrexitDate)
#print("Finished! That took..." + str(datetime.datetime.now() - start))

df, df_links = extract_info_from_zip(zip_filepath)
df.to_csv(output_dir + r'\teleported-references.csv', sep=',',index=False) 
log('Teleported references exported to: ' + output_dir + r'\teleported-references.csv')
df_links.to_csv(links_dir + r'\psl-links.csv', sep=',',index=False) 
log('PSL links exported to: ' + links_dir + r'\psl-links.csv')

log("Finished! That took..." + str(datetime.datetime.now() - start))