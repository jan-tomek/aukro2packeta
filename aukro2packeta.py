import argparse
import re
import imaplib
import email.policy
import random
import string
import time
import base64
import quopri
import requests
import webbrowser
import importlib
import PySimpleGUI as sg
import gettext
import pypdfium2 as pdfium

import config
import packeta

import os, tempfile, subprocess

_ = gettext.gettext


parser = argparse.ArgumentParser(description = _('Creates Packeta package from Aukro email.'))
parser.add_argument("--conf", required = False, help = _('Config file name.'))
parser.add_argument("--pb", required = False, help = _('Only print existing barcode provided as this parameter.'))
parser.add_argument("--verbose", action = 'store_true', help = _('Shows details about the results of running script.'))
parser.add_argument("--log", required = False, help = _('Put all data into files into specified directory.'))

args = parser.parse_args()

if args.conf: conf_name = args.conf
else: conf_name = 'config'
conf = importlib.import_module(conf_name)

api_password = conf.PACKETA_API_PASSWORD
printBarcodeOnly = args.pb
verbose = args.verbose
log = args.log
if log: print(_('Detailed logging into %s directory.' % log))

if printBarcodeOnly: 
    barcode = printBarcodeOnly
    print(_('Printing barcode "%s".' % barcode))

    pckta = packeta.Packeta(verbose, log, api_password)

    pckta.download_convert_print_barcode(barcode)

    exit()

addressId = '0'
number = conf.DEF_NUMBER
name = conf.DEF_NAME
surname = conf.DEF_SURNAME
tgt_email = conf.DEF_EMAIL
tgt_phone = conf.DEF_PHONE
eshop = conf.ESHOP
length = conf.DEF_LENGTH
width = conf.DEF_WIDTH
height = conf.DEF_HEIGHT
weight = conf.DEF_WEIGHT
value = conf.DEF_VALUE

#<street>Českomoravská</street>
#<houseNumber>2408/1a</houseNumber>
#<city>Praha</city>
#<zip>19000</zip>


random_id = ''.join(random.choice(string.digits + string.ascii_letters) for i in range(16))
if verbose: print(_('Random id: %s') % random_id)  

# connect to host using SSL and login to server
if verbose: print(_('IMAP host: %s, user: %s, folder: %s') % (conf.IMAP_HOST, conf.IMAP_USER, conf.IMAP_FOLDER))
imap = imaplib.IMAP4_SSL(conf.IMAP_HOST)
imap.login(conf.IMAP_USER, conf.IMAP_PASS)

# use folder
imap.select(conf.IMAP_FOLDER, readonly=True)

print(_("Looking for unread aukro emails..."))
# search unread email with "dopravy a platby" in subject
typ, data = imap.search("utf-8", 'UNSEEN FROM "oznameni@aukro.cz"')

address_url = None

# take first email
for num in data[0].split():

    typ, emaildata = imap.fetch(num, '(RFC822)')

    if log: open(log+'/'+str(random_id)+'_email_raw.txt', "wb").write(emaildata[0][1])
    msg = email.message_from_bytes(emaildata[0][1], policy = email.policy.default)

    # get email text
    msgtxt = str(quopri.decodestring(msg.as_string()).decode('utf8',errors='ignore'))
    if "Odešlete_prosím_zbož" not in msgtxt: continue
    if log: open(log+'/'+str(random_id)+'_email_txt.txt', "w").write(msgtxt)
    if verbose: print("="*70+'\n'+str(msg.items())+'\n'+"="*70)

    # find subject
    for item in msg.items():
        if item[0] == 'Subject':
            sub = email.header.decode_header(item[1])[0][0]
            print(_('Email Subject is - "%s"') % sub)
        if item[0] == 'Reply-To':
            reply = email.header.decode_header(item[1])[0][0]
            print(_('Email Reply-To is - "%s"') %  reply )
   
    textCore = msgtxt
    if verbose: print("#"*70+'\n'+str(textCore)+'\n'+"#"*70)

    service = ''
    if 'Zásilkovna ČR na výdejní místo' in textCore: service = 'CR_na_vydejni_misto'
    if 'Zásilkovna ČR na adresu' in textCore: service = 'CR_na_adresu'
    if 'Zásilkovna SK na adresu' in textCore: service = 'SK_na_adresu'

    if service == '':
        print(_("No supported Packeta service."))
        exit(1)
    print(_('Service - "%s"') % service)

    tags = re.findall("<[^>]+>",textCore)
    for tag in tags:
        textCore = textCore.replace(tag,'')
    textCore = textCore.replace('=\r\n','')
    textCore = textCore[textCore.find('body {margin: 0; padding: 0;}')+30:]
    textCore = textCore[0:textCore.find('podpora Aukro')]
    textCore = textCore.replace('    ','')

    if log: open(log+'/'+str(random_id)+'_core_text.txt', "w").write(textCore)

    # extract info from email
    number = re.search('%s(.*?)%s' % ('color: #767676;">', '</a>'), msgtxt).group(1).strip()
    if 'misto' in service:
        address_url = 'https://www.zasilkovna.cz/pobocky/' + re.search('%s(.*?)%s' % ('https://www.zasilkovna.cz/pobocky/', '\\?sleId='), msgtxt).group(1).strip()
        name_surname = re.search('%s(.*?)%s' % ('Jm.no a p..jmen.:', 'V.dejn. m.sto:'), textCore).group(1).strip()
    if 'adresu' in service:
        address_url = None
        name_surname = re.search('%s(.*?)%s' % ('Jm.no a p..jmen.:', 'Adresa:'), textCore).group(1).strip()
        tgt_address = re.search('%s(.*?)%s' % ('Adresa:', 'E&#8209;mail:'), textCore).group(1).strip()
        tgt_street = tgt_address.split('  ')[0].strip()
        tgt_rest = tgt_address.split('  ')[1].strip()
        tgt_zip = tgt_rest.split(',')[0].strip()
        tgt_city = tgt_rest.split(',')[1].strip()
        tgt_cntry = tgt_rest.split(',')[2].strip()
        print(_('Address is: %s, %s, %s, %s') % (tgt_street,tgt_zip,tgt_city,tgt_cntry))
    name = name_surname.split(' ')[0]
    surname = name_surname.replace(name, '').strip()
    tgt_email = re.search('%s(.*?)%s' % ('E&#8209;mail:', 'Telefon:'), textCore).group(1).strip()
    tgt_phone = re.search('%s(.*?)%s' % ('Telefon:', 'ZOBRAZIT VÍCE'), textCore).group(1).strip()


    # When not exists or old branch list download it
    if os.path.exists('branch.csv'): file_time = os.path.getmtime('branch.csv')
    else: file_time = 0
    if not os.path.exists('branch.csv') or os.path.getsize('branch.csv') == 0 or ((time.time() - file_time) / 3600 > 24 * 30):
        print(_('Downloading branch.csv file from Packeta.'))
        branchCsvFileReq = requests.get('https://www.zasilkovna.cz/api/v4/' + api_password + '/branch.csv', allow_redirects = True)
        branchCsvFile = open('branch.csv', 'wb')
        branchCsvFile.write(branchCsvFileReq.content)
        branchCsvFile.close()

    # Find info about target branch
    if address_url:
        print(_('Getting Packeta branch id from branch URL - "%s"') % address_url)
        with open('branch.csv', 'r', encoding = 'utf-8-sig') as file:
            addressId = ''
            for line in file:
                if address_url in line:
                    if verbose: print(line.rstrip())
                    addressId = line.split(';')[0].replace('"', '')
                    print(_('Packeta branch is: %s') % addressId)
                    break
            if addressId == '': print(_('Packeta branch id not found.'))
    
imap.close()

# Create dialog box
layout = [
    [sg.Text('')],
    [sg.Text(_('Auction Number'),    size =(15, 1)), sg.InputText(number,    size =(20, 1))],
    [sg.Text('')],
]
if 'misto' in service:
    layout.append([
    [sg.Text(_('Packeta Branch Id'),    size =(15, 1)), sg.InputText(addressId,    size =(20, 1))],
    [sg.Text(_('Packeta Branch Link'), text_color="#0000EE", font=(None, 10), enable_events=True, key="-LINK-")],
        ])
if 'adresu' in service:
    layout.append([
    [sg.Text(_('Street'), size=(15, 1)), sg.InputText(tgt_street, size=(50, 1))],
    [sg.Text(_('City'), size=(15, 1)), sg.InputText(tgt_city, size=(25, 1))],
    [sg.Text(_('ZIP'), size=(15, 1)), sg.InputText(tgt_zip, size=(25, 1))],
    [sg.Text(_('Country'), size=(15, 1)), sg.InputText(tgt_cntry, size=(25, 1))],
    ])
layout.append([
    [sg.Text('')],
    [sg.Text(_('Name'),    size =(15, 1)), sg.InputText(name,    size =(25, 1))],
    [sg.Text(_('Surname'), size =(15, 1)), sg.InputText(surname, size =(25, 1))],
    [sg.Text(_('Email'),   size =(15, 1)), sg.InputText(tgt_email, size =(25, 1))],
    [sg.Text(_('Phone'),   size =(15, 1)), sg.InputText(tgt_phone, size =(25, 1))],
    [sg.Text(_('eshop'),   size =(15, 1)), sg.InputText(eshop,   size =(25, 1))],
    [sg.Text('')],
    [sg.Text(_('Package Length'), size =(15, 1)), sg.InputText(length, size =(15, 1)), sg.Text(_('cm'))],
    [sg.Text(_('Package Width'), size =(15, 1)), sg.InputText(width, size =(15, 1)), sg.Text(_('cm'))],
    [sg.Text(_('Package Height'), size =(15, 1)), sg.InputText(height, size =(15, 1)), sg.Text(_('cm'))],
    [sg.Text('')],
    [sg.Text(_('Package Weight'), size =(15, 1)), sg.InputText(weight, size =(15, 1)), sg.Text(_('kg'))],
    [sg.Text('')],
    [sg.Text(_('Package Value'), size =(15, 1)), sg.InputText(value, size =(15, 1)), sg.Text(_('CZK'))],

    [sg.Text('')],
    [sg.Submit('Create Packeta Package'), sg.Cancel()],
])
  
window = sg.Window('Packeta from Aukro.', layout, finalize=True)
if 'misto' in service:
    window["-LINK-"].set_cursor("hand2")
    window["-LINK-"].Widget.bind("<Enter>", lambda _: window["-LINK-"].update(font=(None, 10, "underline")))
    window["-LINK-"].Widget.bind("<Leave>", lambda _: window["-LINK-"].update(font=(None, 10)))

while True:
    event, values = window.read()

    if event == 'Cancel' or event is None or event == sg.WIN_CLOSED:
        print(_('Dialog canceled.'))
        print(_('Exiting.'))
        exit(1)

    elif event == "-LINK-":
        if address_url == '' or address_url is None: address_url = 'https://www.zasilkovna.cz/pobocky/' + values[0]
        webbrowser.open(address_url)

    elif event == 'Create Packeta Package':
        break

window.close()
 
if 'misto' in service:
    number = int(values[0])
    addressId = int(values[1])
    name = values[2]
    surname = values[3]
    tgt_email = values[4]
    tgt_phone = values[5].replace('(', '').replace(')', '')
    eshop = values[6]
    length = int(values[7]) * 10
    width = int(values[8]) * 10
    height = int(values[9]) * 10
    weight = float(values[10].replace(',','.'))
    value = int(values[11])

if 'adresu' in service:
    number = int(values[0])
    tgt_street = values[1]
    tgt_city = values[2]
    tgt_zip = values[3]
    tgt_cntry = values[4]
    name = values[5]
    surname = values[6]
    tgt_email = values[7]
    tgt_phone = values[8].replace('(', '').replace(')', '')
    eshop = values[9]
    length = int(values[10]) * 10
    width = int(values[11]) * 10
    height = int(values[12]) * 10
    weight = float(values[13].replace(',','.'))
    value = int(values[14])

xml_vars = { 'apiPassword':api_password, 'addressId':addressId, 'number':number, 'name':name, 'surname':surname, 'email':tgt_email, 'phone':tgt_phone, 'eshop':eshop, 'length':length, 'width':width, 'height':height, 'weight':weight, 'value':value}
print(xml_vars)
exit(1)

print(_('Dialog data validated.'))
if verbose: print("="*70+'\n'+str(xml_vars)+'\n'+"="*70)

pckta = packeta.Packeta(verbose, log, api_password)

barcode = pckta.create_package(number,xml_vars)

if log:
    os.rename(log+'/'+str(random_id)+'_email_raw.txt',log+'/'+str(barcode)+'_email_raw.txt')
    os.rename(log+'/'+str(random_id)+'_email_txt.txt',log+'/'+str(barcode)+'_email_txt.txt')
    os.rename(log+'/'+str(random_id)+'_core_text.txt',log+'/'+str(barcode)+'_core_text.txt')

print(_('Printing barcode "%s".' % barcode))

pckta.print_cmd_line = conf.PRINT_CMD_LINE
pckta.print_ext_arg = conf.PRINT_EXT_ARG

pckta.download_convert_print_barcode(barcode)

