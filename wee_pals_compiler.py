import argparse
import random
from glob import glob
import docx
from PyPDF2 import PdfFileReader, PdfFileMerger
from docx2pdf import convert
import os



def convert_docx_to_pdf(file):

    '''
    Convert word file (.docx) to pdf
    Takes in file name

    '''
    wdFormatPDF = 17

    # get abs path of filename
    in_file = os.path.abspath(file)
    # get leading folders and file name
    path, fname = in_file.split('word_docs/')
    # get file name without extension
    fname = fname.split('.docx')[0]
    # concatenate paths and new filename to create abs path for out file
    out_file = path + 'pdfs/' + fname + '.pdf'
    convert(in_file, out_file)



def batch_convert_docs(folder):

    counter = 0

    docFiles = glob(folder + '/*.docx')

    for f in docFiles:

        convert_docx_to_pdf(f)

        counter +=1

    print (counter, 'docs converted')


def compile_pdfs(pdf_folder, year, pdict=None):

    '''

    Merge folder of pdfs into one pdf file

    '''
    
    pdf_list = glob(pdf_folder + '/*.pdf')

    # sort alphabetically
    pdfs = sorted(pdf_list)

    # open merge obj
    merger = PdfFileMerger()

    # loop through pdfs
    for pdf in pdfs:

        # add each pdf to merger object
        merger.append(open(pdf, 'rb'))

    # write all pdfs into one file
    with open('../wee_pals_' + year + '/wee_pals_all_' + year + '.pdf', 'wb') as outPdf:
        merger.write(outPdf)
        outPdf.close()




## takes arguments from command line
if __name__ == '__main__':


    # takes in year argument
    parser = argparse.ArgumentParser()
    parser.add_argument('year', help='what year is it?')
    parser.add_argument('--func', help='function to run; either convert (to convert word docs to pdf) or compile (to compile pdfs)')
    args = parser.parse_args()

    # converts year to str
    year = str(args.year)

    # gets folders for word and pdf docs
    wordFolder = '../wee_pals_' + year + '/word_docs'
    pdfFolder = '../wee_pals_' + year + '/pdfs'

    # if include the optional function argument
    if args.func:
        # either do convert by itself
        if args.func == 'convert':

            batch_convert_docs(wordFolder)

        # or compile by itself
        elif args.func == 'compile':
            
            compile_pdfs(pdfFolder, year)
            
            print('pdfs compiled')
        
    # if no func argument given
    # convert and compile
    else:

        # convert all word docs to pdf
        batch_convert_docs(wordFolder)

        # merge all pdfs
        compile_pdfs(pdfFolder, year)
        print('pdfs compiled')
