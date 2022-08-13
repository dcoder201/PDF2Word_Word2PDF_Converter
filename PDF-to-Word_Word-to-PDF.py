# import docx2pdf
# from pdf2docx import parse
# from docx2pdf import convert

# aspose module integrated with various new features than pdf2docx and docx2pdf
import aspose.words as aw

# for file format checking
import re

# prompting user to select pdf-to-word or word-to-pdf conversion
user_choice = input('PDF-to-WORD  or  WORD-to-PDF (select P2W or W2P)?: ')

# if user selects pdf-to-word
if user_choice.lower() == 'p2w':
    # prompting user to input pdf file name or path to convert
    pdf_file_name = input('Provide PDF Filename or path to convert: ')
    # saving file using aspose for processing
    pdf = aw.Document(pdf_file_name)
    # prompting user to input desired output file name
    docx_output = input('Desired output file name:  ')
    # if user enters  filename with .doc it will do nothing
    if re.search('.doc',docx_output):
        pass
    # if user enters filename without extension , will add proper extension to it using string addition
    else:
        docx_output += '.doc'
    # saving document using user desired output name
    pdf.save(docx_output)

# if user selects word-to-pdf
elif user_choice.lower() == 'w2p':
    # prompting user to input word file name or path to convert
    docx_input = input('Provide DOCX Filename or path to convert: ')
    # saving file using aspose for processing
    doc = aw.Document(docx_input)
    # prompting user to input desired output file name
    pdf_output = input('Desired output file name:  ')
    # if user enters  filename with .pdf it will do nothing
    if re.search('.pdf',pdf_output):
        pass
    # if user enters filename without extension, will add proper extension to it using string addition
    else:
        pdf_output += '.pdf'
    # saving document using user desired output name
    doc.save(pdf_output)
else:
    print('Please choose an option mentioned in option field!')
