#!/usr/bin/python
# -*- coding: utf-8 -*-

import epub
import progressbar
import sys
import os
import glob
import argparse
from docx import Document
from docx.shared import Inches
from html.parser import HTMLParser

class data_parser(HTMLParser):

    """Docstring for data_parser. """

    def __init__(self):
        """TODO: to be defined1. """
        HTMLParser.__init__(self)
        self.data = []
        self.processing = None
        self.paragraph = ''
        
    def handle_starttag(self, tag, attr):
        if tag == 'p':
            self.processing = 'paragraph'

    def handle_data(self, data):
        if self.processing:
            self.paragraph += data

    def handle_endtag(self, tag):
        if self.processing:
            if tag == 'p':
                self.data.append(self.paragraph)
                self.paragraph = ''
                self.processing = None

        
def get_content(epub_file):
    """TODO: Docstring for get_content.

    :epub_file: TODO
    :returns: TODO

    """
    list_data = []
    with epub.open_epub(epub_file) as efile:
        print('reading data from file {}...'.format(epub_file))
        for item_id, linear in efile.opf.spine.itemrefs:
            item = efile.get_item(item_id)
            data = efile.read_item(item)
            data = data.decode('utf-8')
            list_data.append(data)
    epub_file = epub_file.replace('.epub', '')
    return list_data, epub_file
    
def turn_to_doc(l_data, doc_filename, p_count):
    """TODO: Docstring for get_content.

    :l_data: TODO
    :epub_file: TODO
    :returns: TODO

    """
    try:
        template_file = os.path.join(sys._MEIPASS, 'templates/default.docx') 
    except AttributeError:
        template_file = None

    new_document = Document(template_file)
    bar_count = 0
    with progressbar.ProgressBar(max_value=p_count, redirect_stdout=True) as bar:
        for data in l_data:
            for item in data:
                new_document.add_paragraph(item)
                bar_count += 1
                bar.update(bar_count)
    new_document.save(os.path.abspath(os.path.join('docx/',doc_filename)))
    print('Success!File saved as {}\n'.format(os.path.abspath(os.path.join('docx/',doc_filename))))
    return 0

def main():
    """TODO: Docstring for get_content.

    :path: TODO
    
    """
    parser = argparse.ArgumentParser(description='Translate epub to docx')
    parser.add_argument('path', help='the path where the epub files are') 
    args = vars(parser.parse_args())
    always_write = False
    always_skip = False
    final_data_list = []

    print('change directory to %s' %args['path'])
    os.chdir(args['path'])
    try:
        os.mkdir('docx')
    except FileExistsError:
        print('WARNING:Ducument docx exists, files may be overwrote!\n')

    for epub_file in glob.glob('*.epub'):
        p_count = 0
        list_data,filename = get_content(epub_file)
        doc_filename = filename + '.docx'
        parser = data_parser()
        for data in list_data:
            parser.feed(data)
            final_data = parser.data
            data_copy = final_data
            for i, item in enumerate(final_data):
                if '\n' in item:
                    data_copy.pop(i)
            final_data_list.append(data_copy)
        for item in final_data_list:
            p_count += len(item)
        if os.path.exists(os.path.abspath(os.path.join('docx/',doc_filename))) and not always_write:
            if always_skip:
                continue
            print("'{}' already exists. Overwrite? [y/N/all/none] ".format(os.path.abspath(os.path.join('docx/',doc_filename)))),
            response = sys.stdin.readline().strip().lower()
            if (response == 'all'):
                always_write = True
                turn_to_doc(final_data_list, doc_filename, p_count)
            elif (response == 'y' or response == "yes" ):
                turn_to_doc(final_data_list, doc_filename, p_count)
            elif (response == 'none'):
                always_skip = True
                continue
            else:
                continue
        else:
            turn_to_doc(final_data_list, doc_filename, p_count)
        final_data_list = []

if (__name__ == "__main__"):
    sys.exit(main())

