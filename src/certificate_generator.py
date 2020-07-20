"""
Created on Sun Jul 19 01:19:20 2020

@author: Aarti
"""


# Refer to artifacts folder to know the file format to be used in this program
# Recommendation: program works best when used for creating 50 certificates at a time
# Warning: be mindful of creating both the .txt file in same order

from docx import Document
import os
import itertools
from docx2pdf import convert
from docx.oxml.xmlchemy import OxmlElement
from docx.oxml.shared import qn

# Variable Declaration
main_dir = "E:\MTech\CyberSecurityQuiz\Certificates"    # path to dir containing names.txt, uni.txt & template.docx
cert_dir = "E:\MTech\CyberSecurityQuiz\Certificates\pdf_certificates"  # path to directory to store certificates
name_list = "names.txt"                # .txt file containing full names of participants
uni_name_list = "uni.txt"              # .txt file containing full name of uni for each participant
cert_template = 'Participant Certificate Template.docx'  # .docx template file name

os.chdir(main_dir)

# reading content of name_list and uni_name_list
with open(name_list) as f:
    names = f.read()
with open(uni_name_list) as f1:
    uni = f1.read()

# creating list of all participants name and university names, separated by new line
list_of_names = names.split('\n')
list_of_uni_names = uni.split('\n')

# running for loop for each participant name and corresponding university name
for (a, b) in itertools.zip_longest(list_of_names, list_of_uni_names):
    os.chdir(main_dir)
    document = Document(cert_template)
    for p in document.paragraphs:
        inline = p.runs
        # Loop added to work with runs (strings with same style)
        for i in range(len(inline)):
            if 'Meghna Sharma' in inline[i].text:
                text = inline[i].text.replace('Meghna Sharma', a)
                inline[i].text = text
            if 'USICT' in inline[i].text:
                text = inline[i].text.replace('USICT', b)
                inline[i].text = text
    # Adding border to the certificate
    sec_pr = document.sections[0]._sectPr # get the section properties el
    # create new borders el
    pg_borders = OxmlElement('w:pgBorders')
    # specifies how the relative positioning of the borders should be calculated
    pg_borders.set(qn('w:offsetFrom'), 'page')
    for border_name in ('top', 'left', 'bottom', 'right',): # set all borders
        border_el = OxmlElement(f'w:{border_name}')
        border_el.set(qn('w:val'), 'single') # a single line
        border_el.set(qn('w:sz'), '32') # for meaning of  remaining attrs please look docs
        border_el.set(qn('w:space'), '24')
        border_el.set(qn('w:color'), '1A5276')
        pg_borders.append(border_el) # register single border to border el
    sec_pr.append(pg_borders) # apply border changes to section

    # changing directory location to certificate directory
    os.chdir(cert_dir)
    doc_file_name = a+'.docx'
    document.save(doc_file_name)           # saving certificate.docx file to the cert_dir
    convert(doc_file_name)                 # converting that certificate.docx to certificate.pdf
print("end of code!")

