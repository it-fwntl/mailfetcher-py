# pip install python-dotenv, Imbox, PyPDF2, pycups, smbprotocol, graypy
import logging
import graypy
from dotenv import load_dotenv
load_dotenv()
import os
from imbox import Imbox
from PyPDF2 import PdfReader
import cups
import re
import traceback
import xml.etree.ElementTree as ET
import smbclient.shutil
from logging.handlers import TimedRotatingFileHandler

MAIL_VALIDATION_FROM=os.environ.get("MAIL_VALIDATION_FROM")
MAIL_VALIDATION_SUBJECT=os.environ.get("MAIL_VALIDATION_SUBJECT")
MAIL_ACCOUNT_HOST = os.environ.get("MAIL_ACCOUNT_HOST")
MAIL_ACCOUNT_USER = os.environ.get("MAIL_ACCOUNT_USER")
MAIL_ACCOUNT_PASSWORD = os.environ.get("MAIL_ACCOUNT_PASSWORD")
PRINT_CUPS_NAME = os.environ.get("PRINT_CUPS_NAME")
PRINT_ALARM_MAIL = os.environ.get("PRINT_ALARM_MAIL") == "true"
PRINT_ALARM_MAIL_AMOUNT = os.environ.get("PRINT_ALARM_MAIL_AMOUNT")
PRINT_CLOSING_MAIL = os.environ.get("PRINT_CLOSING_MAIL") == "true"
PRINT_CLOSING_MAIL_AMOUNT = os.environ.get("PRINT_CLOSING_MAIL_AMOUNT")
GRAYLOG_HOST = os.environ.get("GRAYLOG_HOST")
SMB_HOST = os.environ.get("SMB_HOST")
SMB_USERNAME = os.environ.get("SMB_USERNAME")
SMB_PASSWORD = os.environ.get("SMB_PASSWORD")
SMB_FOLDER_NAME = os.environ.get("SMB_FOLDER_NAME")
VEHICLES = os.environ.get("VEHICLES").split(";")
IS_READONLY_MODE = os.environ.get("IS_READONLY_MODE") == "1"

FOLDER_DIR = os.path.dirname(os.path.realpath(__file__))
DOWNLOAD_FOLDER = f"{FOLDER_DIR}/temp"

# Takes an xml file as input. Outputs ElementTree and element
def load_xml(name):
    tree = ET.parse(name)
    root = tree.getroot()
    return tree, root

# Print file on defined printer
def print_pdf(file, amount):
    logging.info(f'Print attachment to {amount}x on {PRINT_CUPS_NAME} => {file}')
    conn = cups.Connection()
    conn.printFile(PRINT_CUPS_NAME, file, " ", {"copies": amount})

# Load existed remote xml file or create a new based on template
def load_remote_xml(number):
    path_to_xmls = fr"\\{SMB_HOST}\data\fwntl\{SMB_FOLDER_NAME}\Einsatzdokumentation\Datenimport"
    file_path = f"einsatzdaten_{number}.xml"
    smbclient.register_session(SMB_HOST, username=SMB_USERNAME, password=SMB_PASSWORD)
    # If there is no file (first from alarm mail), copy template to server
    if not file_path in smbclient.listdir(path_to_xmls):
        smbclient.shutil.copyfile(f"{FOLDER_DIR}/einsatzdaten_vorlage.xml", fr"{path_to_xmls}\{file_path}")
    # Download current file to store new informations
    smbclient.shutil.copyfile(fr"{path_to_xmls}\{file_path}", fr"{DOWNLOAD_FOLDER}/{file_path}")
    smbclient.delete_session(SMB_HOST)
    return fr"{DOWNLOAD_FOLDER}/{file_path}"

def save_xml_remote(file):
    path_to_xmls = fr"\\{SMB_HOST}\data\fwntl\{SMB_FOLDER_NAME}\Einsatzdokumentation\Datenimport"
    smbclient.register_session(SMB_HOST, username=SMB_USERNAME, password=SMB_PASSWORD)
    smbclient.shutil.copyfile(fr"{DOWNLOAD_FOLDER}/{file}", fr"{path_to_xmls}\{file}")
    smbclient.delete_session(SMB_HOST)

# Create logging instance
logger = logging.getLogger()
logger.setLevel(logging.INFO)
handler = graypy.GELFTCPHandler(GRAYLOG_HOST, 12201)
logger.addHandler(handler)

# Create temp folder to store mail attachments
if not os.path.isdir(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# Setup mail client
logger.info('Start fetching mails...')
mail = Imbox(MAIL_ACCOUNT_HOST, username=MAIL_ACCOUNT_USER, password=MAIL_ACCOUNT_PASSWORD, ssl=True, ssl_context=None, starttls=False)
messages = mail.messages(unread=True, sent_from=MAIL_VALIDATION_FROM, subject=MAIL_VALIDATION_SUBJECT)

# Load unread mails an store attachments ended with .pdf
for (uid, message) in messages:
    logger.info(f'Mark mail as seen {uid}')
    if IS_READONLY_MODE == False: mail.mark_seen(uid)
    for idx, attachment in enumerate(message.attachments):
        try:
            att_fn = attachment.get('filename')
            download_path = f"{DOWNLOAD_FOLDER}/{att_fn}"
            if ".pdf" in download_path:
                with open(download_path, "wb") as fp:
                    fp.write(attachment.get('content').read())
                    logger.info(f'Store new attachment {att_fn}')
        except:
            logger.error(traceback.print_exc())
mail.logout()

try:
    for filename in os.listdir(DOWNLOAD_FOLDER):
        logger.info(f'Run script for file {filename}')
        if ".pdf" in filename:
            is_alarm_mail = False
            is_closing_mail = False
            if "Alarmdruck" in filename: is_alarm_mail = True
            elif "Einsatzende" in filename:  is_closing_mail = True

            # Print pdf to local printer
            if IS_READONLY_MODE == False and PRINT_ALARM_MAIL and is_alarm_mail: print_pdf(f"{DOWNLOAD_FOLDER}/{filename}", PRINT_ALARM_MAIL_AMOUNT)
            if IS_READONLY_MODE == False and PRINT_CLOSING_MAIL and is_closing_mail: print_pdf(f"{DOWNLOAD_FOLDER}/{filename}", PRINT_CLOSING_MAIL_AMOUNT)
            
            # Convert pdf to text
            reader = PdfReader(f"{DOWNLOAD_FOLDER}/{filename}") 
            text = reader.pages[0].extract_text()
            logger.info(text)
            
            # Gather and store informations from alarm mail
            nummer = re.search("(Alarmdruck|Abschlussbericht)\s+([\d]+)", text).group(2)
            print(nummer)
            xml_file = load_remote_xml(nummer)
            tree, root = load_xml(xml_file)
            informations = root.find('informationen')

            informations.find('einsatznummer').text = nummer

            stichwort = re.search("Stichwort\s(\w+\s\w+)", text).group(1)
            informations.find('stichwort').text = stichwort

            ortsteil = re.search("Ortsteil\s(.+)", text).group(1)
            adresse = ''
            if is_alarm_mail: adresse = re.search("Straße\s(.+(?=Alarmdruck)|.+)", text).group(1)
            elif is_closing_mail: adresse = re.search("Straße\s(.+)", text).group(1)
            if "A61" in adresse: informations.find('adresse').text = adresse
            else: informations.find('adresse').text = f"{adresse}, {ortsteil}"

            meldender = re.search("Meldender\s([\w, -]+)", text).group(1)
            informations.find('meldender').text = meldender

            if is_closing_mail:
                date = re.search("alarmiert\s([\d.]+)", text).group(1)
                informations.find('alarmierungsdatum').text = date
                time = re.search("alarmiert\s[\d.]+\s([\d:]+)", text).group(1)
                informations.find('alarmierungszeit').text = time

                for index, vehicle in enumerate(VEHICLES):
                    result = re.search(fr"{vehicle}\n?(\([/\d]+\)|\d|-|\s)([- \d*:]+)", text)
                    if (result):
                        time_row = result.group(2).split(" ")
                        root.find('sds').find(f"fzg{index+1}").find('einsatzUebernommen').text = time_row[2].replace("*", "").replace("--:--:--","")
                        root.find('sds').find(f"fzg{index+1}").find('ankunftAmEinsatzort').text = time_row[3].replace("*", "").replace("--:--:--","")
                        root.find('sds').find(f"fzg{index+1}").find('einsatzbereitUeberFunk').text = time_row[7].replace("*", "").replace("--:--:--","")
                        root.find('sds').find(f"fzg{index+1}").find('einsatzbereitAufWache').text = time_row[8].replace("*", "").replace("--:--:--","")
            
            tree.write(f"{DOWNLOAD_FOLDER}/einsatzdaten_{nummer}.xml", encoding='utf-8', xml_declaration=True)
            logger.info(ET.tostring(root, encoding='unicode'))
            if IS_READONLY_MODE == False: save_xml_remote(f"einsatzdaten_{nummer}.xml")
            os.remove(f"{DOWNLOAD_FOLDER}/einsatzdaten_{nummer}.xml")
            os.remove(f"{DOWNLOAD_FOLDER}/{filename}")
except Exception as e:
    os.remove(f"{DOWNLOAD_FOLDER}/{filename}")
    os.remove(f"{DOWNLOAD_FOLDER}/einsatzdaten_{nummer}.xml")
    logging.error('Error at %s', 'Parse new xml file', exc_info=e)