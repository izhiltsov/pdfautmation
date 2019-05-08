# Copyright 2019 Ilya Zhiltsov izhiltsov@gmail.com
# Licensed under the Apache License, Version 2.0
# http://www.apache.org/licenses/LICENSE-2.0

"""

1. Import a group of files from a folder
2. Search on a user-provided string.
3. Extract each page of the documents where there is a hit on the search key and
create a separate PDF file for each page where there is a hit.
4. Combine all of the extracted pages into a single PDF file.



"""

import sys
import os

from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from progressbar import ProgressBar


def splitpages(filename, prefix):
    pdf = PdfFileReader(filename, strict=False)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        output_filename = f'{prefix}{page + 1}_{filename}.tmp'

        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)


def page_merger(filenames, output_filename):
    pdf_merger = PdfFileMerger()
    for filename in filenames:
        pdf_merger.append(filename)

    with open(output_filename, 'wb') as f:
        pdf_merger.write(f)


def extract_text(filename):
    with open(filename, 'rb') as f:
        pdf = PdfFileReader(f, strict=False)
        page = pdf.getPage(0)
        text = ''.join(page.extractText().split('\n')).lower()

        return text


def get_files(path, prefix=None):
    if prefix == None:
        filenames = [filename for filename in os.listdir(path) if filename.endswith('.pdf')]
    else:
        filenames = [filename for filename in os.listdir(path) if filename.startswith(prefix)]

    return filenames

def deltemfiles(path, prefix):

    filelist = [ f for f in os.listdir(path) if f.startswith(prefix) ]
    for f in filelist:
        os.remove(os.path.join(path, f))

def main():


    path = '.'
    prefix = 'tmp_page_'
    output_filename = 'result.pdf'  

    # Ask User to input string
    while True:
        user_string = input('Enter string: ')
        if user_string:
            break
       
    filenames = get_files(path)

    print(f'Total pdf files in directory: {len(filenames)}')

    print('Splitting files to pages...')

    pbar1 = ProgressBar()

    for filename in pbar1(filenames):
        splitpages(filename, prefix)

    pages = get_files(path, prefix)
    print('Done.')

    print(f'Total pages to analyse: {len(pages)}')
    print('Please wait...')

    results = []
    counter = 0

    pbar2 = ProgressBar()
    for page in pbar2(pages):
        if user_string.lower() in extract_text(page):
            results.append(page)
            counter += 1

    if counter > 0:
        print(f'We found {counter} pages that contain string: "{user_string}"')
        print("Let's try to create the final PDF file... ")
        page_merger(results, output_filename)
        print('PDF file has created...')
    else:
        print('Not found')

    print('Cleaning Up...')

    deltemfiles(path, prefix)


    print('Done!')




if __name__ == '__main__':
    main()