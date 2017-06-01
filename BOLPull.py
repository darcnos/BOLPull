# -*- coding: utf-8 -*-

import requests, base64, time, os, shutil, glob, csv
from subprocess import call
import smtplib, configparser, ftplib
from datetime import datetime
from pytz import timezone

est = timezone('US/Eastern')
#Where am I running from?
dir_path = os.path.dirname(os.path.realpath(__file__))

# Read INI, set which numbers go to what groups
configobject = configparser.ConfigParser()
configobject.read(dir_path + '\\scriptconfig.ini')

#Just some end points
api = 'https://applications.filebound.com/v4/'
login_end = api + 'login?'
docs_end = api + 'documents/'
files_end = api + 'files/'
fbsite = 'fbsite='
url = fbsite + 'https://burriswebdocs.filebound.com'

#Lists which were used in testing or are currently being used
missingbolpo = []
notmissingbolpo = []

missingorder = []
notmissingorder = []

missingdc = []
notmissingdc = []

print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + "Script executing in: " + dir_path)

currenttime = datetime.now(est).strftime("%Y%m%d-%H_%M")


#For this run, this is where the PDFs will live
global image_dir
image_dir = dir_path + '\\images\\' + currenttime + '\\'
spreadsheets = dir_path + '\\data'
if not os.path.exists(spreadsheets):
    os.makedirs(spreadsheets)
spreadsheetpath = spreadsheets



def login():
    u = configobject['WebDocs']['user']
    p = configobject['WebDocs']['pass']
    data = {
        'username': u,
        'password': p
    }
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Logging into WebDocs as "{}"'.format(u))
    login = login_end + url
    r = requests.post(login, data)
    if r.status_code == 200:
        guid = r.json()
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Logged into WebDocs successfully')
        return guid
    else:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Error when logging into WebDocs. Check your connection and try again.')
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Status code: ' + str(r.status_code))





def customquer2(bolpo):
    url = 'https://applications.filebound.com/v3/query/projectId_2/F1_' + bolpo + '/divider_/binaryData?fbsite=https://burriswebdocs.filebound.com' + guid
    r = requests.get(url)
    if r.status_code != 200:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Network connectivity error when querying {}'.format(bolpo))
    if r.status_code == 200:
        data = r.json()
        if data[0]['files']['Collection']:
            print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Document with PO {} found on site'.format(bolpo))
            notmissingbolpo.append(str(bolpo))
            return(data)
        else:
            print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Document with PO {} not found on site'.format(bolpo))



def bolprocess(data, currentorder, currentDC, currentpo):
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Successfully opened TSM lookup record')
    for file in data[0]['files']['Collection']:
        #fileleveldir = dir_path + '\\images\\' + '\\' + currenttime + '\\' + currentDC + '\\' + str(currentorder)
        fileleveldir = dir_path + '\\images\\' + currenttime + '\\' + str(currentorder)
        if not os.path.exists(fileleveldir):
            os.makedirs(fileleveldir)
        os.chdir(fileleveldir)
        doccount = 0
        for i in file['documents']['Collection']:
            doccount += 1
            docId = i['documentId']
            extension = i['extension']
            binaryData = i['binaryData']
            convertedbinaryData = base64.b64decode(binaryData)
            playfile = str(str(docId) + '.' + extension)
            # print(fileleveldir)
            with open(playfile, 'wb') as f:
                f.write(convertedbinaryData)
            print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Downloaded page #' + str(doccount) + ' of PO {}'.format(currentpo))
        #currentdirectory = (os.getcwd())
        #filesinourdir = [f for f in listdir(currentdirectory) if isfile(join(currentdirectory, f))]
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Saved to: '+fileleveldir)
        os.chdir("..")
    try:
        #call(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", "-windowstyle", "hidden", "convert.exe", fileleveldir + "//*", fileleveldir + '.PDF'])
        call(["C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe", "convert.exe", fileleveldir + "//*", fileleveldir + '.PDF'])
        shutil.rmtree(fileleveldir)
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Converted PO {} to PDF succesfully'.format(currentpo))
    except:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Error when converting PO {} to PDF'.format(currentpo))


def newopen(path):
    for file in glob.glob(spreadsheetpath + '\\*.xls'):
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'TMS lookup file {} found at {}'.format(file ,spreadsheetpath))
        newname = file + '.processing'
        newnewname = file
        os.rename(file, newname)
        with open(newname) as csvfile:
            neat = csv.DictReader(csvfile, delimiter='\t')
            for row in neat:
                currentpo = row['PONumber']
                currentorder = row['ord_hdrnumber']
                currentDC = row['DC']
                print('\n' + datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Performing query on PO {} with location code {}'.format(currentpo, currentDC))
                data = customquer2(currentpo)
                try:
                    bolprocess(data, currentorder, currentDC, currentpo)
                except TypeError as notfound:
                    missingbolpo.append(str(currentpo))
                    missingorder.append(str(currentorder))
                    missingdc.append(str(currentDC))
                    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Added PO ' + str(currentpo) + ' to the missing log')
                except TimeoutError as nointernet:
                    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Cannot reach the WebDocs server. Test and confirm your network connection.')
                except ConnectionError as err:
                    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Cannot reach the WebDocs server. Test and confirm your network connection.')
                except ConnectionAbortedError:
                    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Cannot reach the WebDocs server. Test and confirm your network connection.')
        processed = spreadsheets + '\\completed\\' + currenttime + '\\'
        if not os.path.exists(processed):
            os.makedirs(processed)
        os.rename(newname, newnewname)
        shutil.move(newnewname, processed)
        os.chdir(processed)
        if len(missingbolpo) > 0:
            missingfile = open(processed + '\\missing.txt', 'w')
            for item in missingbolpo:
                missingfile.write("%s\n" % item)



def checkNet():
    try:
        r = requests.get('http://www.google.com/')
        ayy = r.raise_for_status()
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + str(ayy))
        return True
    except requests.exceptions.HTTPError as err:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + err)
        return False
    except requests.exceptions.ConnectionError as err:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + err)
        return False



result = checkNet()
if result == True:
    guid = '&guid=' + login()
    newopen(spreadsheetpath)
if result == False:
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + '\nCannot reach the internet. Check your connection and try again.')


#Create some lists to store which group's docs are missing for emailing
sevenone = []
ohone = []
seventhree = []
fourthree = []
nineone = []

#Assemble the missing PO numbers to the location codes they should be emailed to
for i in range(len(missingbolpo)):
    if missingdc[i] == '071':
        sevenone.append(missingbolpo[i])
    elif missingdc[i] == '001':
        ohone.append(missingbolpo[i])
    elif missingdc[i] == '073':
        seventhree.append(missingbolpo[i])
    elif missingdc[i] == '043':
        fourthree.append(missingbolpo[i])
    elif missingdc[i] == '091':
        nineone.append(missingbolpo[i])
    else:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'If you see this, then your input spreadsheet has unrecognized values in the DC column.\n')


#Ready the ini's email configs
subject = configobject['Emails']['subject']
body = configobject['Emails']['body']
email_from = configobject['Emails']['from']
p = configobject['Emails']['pass']



#This block code looks at the length of the missing group arrays
#and if the length of an array is more than 0, it converts the array's content
#to string and then places it inside an email body to sen
print('----')
print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Lookup and retrievals complete.')
print('----')




if len(missingbolpo) > 0:
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Emailing missing PO numbers to email groups')
    try:
        if len(sevenone) > 0:
            email_to = configobject['Emails']['071']
            email_sevenone = ", ".join(sevenone)
            gm = email_session(email_from, email_to)
            gm.send_message(subject, body.format(email_sevenone))
        if len(ohone) > 0:
            email_to = configobject['Emails']['001']
            email_ohone = ", ".join(ohone)
            gm = email_session(email_from, email_to)
            gm.send_message(subject, body.format(email_ohone))
        if len(seventhree) > 0:
            email_to = configobject['Emails']['073']
            email_seventhree = ", ".join(seventhree)
            gm = email_session(email_from, email_to)
            gm.send_message(subject, body.format(email_seventhree))
        if len(fourthree) > 0:
            email_to = configobject['Emails']['043']
            email_fourthree = ", ".join(fourthree)
            gm = email_session(email_from, email_to)
            gm.send_message(subject, body.format(email_fourthree))
        if len(nineone) > 0:
            email_to = configobject['Emails']['091']
            email_nineone = ", ".join(nineone)
            gm = email_session(email_from, email_to)
            gm.send_message(subject, body.format(email_nineone))
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Emails sent sucessfully')
        emailsuccess = True
    except:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'WARNING!')
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") +'Unable to send emails. Check SMTP server, and email credentials.\n')
        emailsuccess = False
else:
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'No emails to send at this time.')




### FTP beyond this point
server = configobject['FTP']['server']
username = configobject['FTP']['username']
password = configobject['FTP']['password']
purgefiles = configobject['housekeeping']['purge_images']

def uploadThis(path):
    files = os.listdir(path)
    os.chdir(path)
    for f in files:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Uploading {} to FTP Server at {} as {}'.format(f, server, username))
        if os.path.isfile(path + r'\{}'.format(f)):
            fh = open(f, 'rb')
            myFTP.storbinary('STOR %s' % f, fh)
            fh.close()
        elif os.path.isdir(path + r'\{}'.format(f)):
            myFTP.mkd(f)
            myFTP.cwd(f)
            uploadThis(path + r'\{}'.format(f))
    myFTP.cwd('..')
    os.chdir('..')
    myFTP.quit()


if not os.path.exists(image_dir):
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'No TMS lookup file to process in: {}'.format(spreadsheetpath))
else:
    downloaded_images = os.listdir(image_dir)
    print('\n' + datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Uploading PDFs from {}'.format(image_dir))
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Connecting to FTP at {} as {}'.format(server, username))
    try:
        myFTP = ftplib.FTP(server, username, password)
        uploadThis(image_dir)
        if purgefiles == 'True':
            print('\n' + datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Downloaded images from this run are set to delete')
            shutil.rmtree(image_dir)
            print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Downloaded images have been removed')
        else:
            print('\n' + datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Downloaded images are set to NOT delete')
    except:
        print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Unable to reach FTP server at {} as {}'.format(server, username))
        exit()

if emailsuccess != True:
    print('\n' + datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Script executed, but emails were NOT sent')
    print(datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Check the email server and the ini and confirm the email account credentials are correct')

else:
    print('\n' + datetime.now(est).strftime("%m/%d/%Y %H:%M:%S - ") + 'Script executed successfully')