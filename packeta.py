import gettext
import requests
import base64
import re
import os
import tempfile
import subprocess
import pypdfium2 as pdfium
       	
class Packeta:

    def __init__(self, verbose, log, api_password):
        self.verbose = verbose
        self.log = log
        self.api_password = api_password

        self.print_cmd_line = "C:\\Program Files\\IrfanView\\i_view64.exe"
        self.print_ext_arg = "/print=Label_printer"

    def create_package(self, number, xml_vars):

        print(xml_vars)

        if xml_vars['type'] == 'misto':
            packeta_pack_req_xml = '''
            <createPacket>
            <apiPassword>{apiPassword}</apiPassword>
              <packetAttributes>
                <addressId>{addressId}</addressId>
                <number>{number}</number>
                <name>{name}</name>
                <surname>{surname}</surname>
                <email>{email}</email>
                <phone>{phone}</phone>
                <eshop>{eshop}</eshop>
                <size>
                  <length>{length}</length>
                  <width>{width}</width>
                  <height>{height}</height>
                </size>
                <weight>{weight}</weight>
                <value>{value}</value>
              </packetAttributes>
            </createPacket>
            '''.format(**xml_vars)

        if xml_vars['type'] == 'adresa':
            packeta_pack_req_xml = '''
            <createPacket>
            <apiPassword>{apiPassword}</apiPassword>
              <packetAttributes>
                <street>{street}</street>
                <houseNumber>{house_number}</houseNumber>
                <city>{city}</city>
                <zip>{zip}</zip>
                <number>{number}</number>
                <name>{name}</name>
                <surname>{surname}</surname>
                <email>{email}</email>
                <phone>{phone}</phone>
                <eshop>{eshop}</eshop>
                <size>
                  <length>{length}</length>
                  <width>{width}</width>
                  <height>{height}</height>
                </size>
                <weight>{weight}</weight>
                <value>{value}</value>
              </packetAttributes>
            </createPacket>
            '''.format(**xml_vars)

        print('XML for Packeta generated.')
        if self.verbose: print("=" * 70 + packeta_pack_req_xml + "=" * 70)
        if self.log: open(self.log + '/' + str(number) + '_rqst.xml', "w").write(packeta_pack_req_xml)

        # sending get request and saving the response as response object
        resp = requests.get(url='https://www.zasilkovna.cz/api/rest', headers={"Content-Type": "text/xml"},
                            data=packeta_pack_req_xml.encode('utf-8'))
        if self.verbose: print("=" * 70 + '\n' + resp.text + "=" * 70)
        if self.log: open(self.log + '/' + str(number) + '_resp.xml', "w").write(resp.text)

        if resp.status_code == 200:
            print('HTTP Post OK')
            barcode = re.search('%s(.*?)%s' % ('<barcode>', '</barcode>'), resp.text).group(1)
            print('Barcode for package: ' + barcode)
            return barcode

        else:
            print('Error sending API request to Packeta.')
            exit(2)

    def download_barcode(self, bar_code, pdf_file_name):

        _ = gettext.gettext

        xml_vars = { 'apiPassword':self.api_password, 'barcode':bar_code}

        packeta_bar_code_req_xml = '''
        <packetLabelPdf>
        <apiPassword>{apiPassword}</apiPassword>
          <packetId>{barcode}</packetId>
          <format>A6 on A6</format>
          <offset>0</offset>
        </packetLabelPdf>\n'''.format(**xml_vars).replace('        ','')

        print(_('XML to get barcode generated.'))
        if self.verbose: print("="*70+packeta_bar_code_req_xml+"="*70)
        if self.log: open(self.log +'/' + str(bar_code) + '_rqst.xml', "w").write(packeta_bar_code_req_xml)
        
        print(_('Downloading pdf barcode file from Packeta.'))
        resp = requests.get(url = 'https://www.zasilkovna.cz/api/rest', headers = {"Content-Type": "text/xml"}, data = packeta_bar_code_req_xml)
        if self.verbose: print("="*70+'\n'+resp.text[0:110]+'\n'+"="*70)  
        if self.log: open(self.log +'/' + str(bar_code) + '_resp.xml', "w").write(resp.text)
        
        print(_('base64 response generate.'))
        base64str = resp.text
        base64str = re.sub('..xml version=.1.0. encoding=.utf.8...','',base64str)
        base64str = re.sub('.response..status.ok..status..result.','',base64str)
        base64str = re.sub('..result...response.','',base64str)
        base64str = re.sub('\n','',base64str)
        if self.log: open(self.log +'/' + str(bar_code) + '.base64', "w").write(base64str)
        
        print(_('Creating pdf barcode into %s file.' % pdf_file_name))
        pdf = base64.urlsafe_b64decode(base64str)
        pdf_file = open(pdf_file_name, "wb")
        pdf_file.write(pdf)
        pdf_file.close()
        if self.log: open(self.log +'/' + str(bar_code) + '.pdf', "wb").write(pdf)
        
    def convert_barcode(self, bar_code, pdf_file_name, png_file_name):

        _ = gettext.gettext

        print(_('Converting "%s" to image file "%s".') % (pdf_file_name, png_file_name))

        pdf = pdfium.PdfDocument(pdf_file_name)
        page = pdf.get_page(0)
        pil_image = page.render(
            scale=8,
            rotation=90,
            crop=(0, 0, 0, 0)
        ).to_pil()
        pdf.close()
        
        width, height = pil_image.size 
        if self.verbose: print('Image size: ' +str(width) + ' x ' + str(height) + ' pixels.')
        pil_image = pil_image.crop((0, 80, width, height - 100))
        pil_image = pil_image.convert('L').point(lambda x : 255 if x > 200 else 0, mode = '1')
        if self.verbose: pil_image.show()
        if self.log: pil_image.save(self.log +'/' + str(bar_code) + '.png')
        
        print(_('Saving image for print as "%s".') % png_file_name)
        pil_image.save(png_file_name)

    def print_barcode(self, bar_code, png_file_name):

        _ = gettext.gettext

        print(_('Printing label for %s from file "%s".') % (bar_code,png_file_name))

        if self.verbose: print(_('Running print command: %s %s %s') % (self.print_cmd_line, png_file_name, self.print_ext_arg))
        subprocess.call([self.print_cmd_line, png_file_name, self.print_ext_arg])

        print('Print done.')

    def download_convert_print_barcode(self, bar_code):

        _ = gettext.gettext

        print(_('Processing "%s".') % bar_code)

        pdf_temp_name = next(tempfile._get_candidate_names()) + '.pdf'
        if self.verbose: print(_('PDF temp file "%s".') % pdf_temp_name)
        png_temp_name = next(tempfile._get_candidate_names()) + '.png'
        if self.verbose: print(_('PNG temp file "%s".') % png_temp_name)

        self.download_barcode(bar_code, pdf_temp_name)
        self.convert_barcode(bar_code, pdf_temp_name, png_temp_name)
        self.print_barcode(bar_code, png_temp_name)

        if self.verbose: print(_('Removing temp files.'))
        os.remove(pdf_temp_name)
        os.remove(png_temp_name)










          