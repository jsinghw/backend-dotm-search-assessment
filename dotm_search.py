#!/usr/bin/env python 3
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "jsingw"

import argparse
import os
import zipfile

DOC_FILENAME = 'word/document.xml'


def search_zip(z, search_text, full_path):
    with z.open(DOC_FILENAME) as doc:
        xml_text = doc.read().decode('utf8')
    text_location = xml_text.find(search_text)
    if text_location >= 0:
        print('Match found in file {}'.format(full_path))
        print('   ...' + xml_text[
            text_location-40:text_location+40] + '...   ')
        return True
    return False


def create_parser():
    parser = argparse.ArgumentParser(description='Description here')
    parser.add_argument('--dir', action='store', default=os.getcwd(),
                        dest='dir', help='Path provided')
    parser.add_argument('searchTerm', action='store', type=str,
                        help='Search term ')
    args = parser.parse_args()
    return(args)


def main():
    args = create_parser()
    search_text = args.searchTerm
    search_path = args.dir

    print("Searching directory {} for dotm files with text '{}' ...".format(
        search_path, search_text))

    file_list = os.listdir(search_path)

    match_count = 0
    search_count = 0

    for file in file_list:
        if file.endswith('.dotm'):
            search_count += 1

        full_path = os.path.join(search_path, file)

        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as z:
                names = z.namelist()
                if DOC_FILENAME in names:
                    if search_zip(z, search_text, full_path):
                        match_count += 1
        else:
            print("Not a zipfile: " + full_path)
    print('Total dom files searched: {}'.format(search_count))
    print('total dotm files matched: {}'.format(match_count))


if __name__ == '__main__':
    main()
