#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Patrick Martin"


import os
import sys
import argparse
import zipfile


def main(args):

    path_to_dotm_files = os.path.join(os.getcwd(), 'dotm_files')

    parser = argparse.ArgumentParser()
    parser.add_argument(
        '--dir', help='Specify the directory to search for .dotm files',
        default=path_to_dotm_files)
    parser.add_argument(
        'text', help='The text to string search .dotm filse for')

    args = parser.parse_args()

    file_list = os.listdir(args.dir)
    file_list = [f for f in file_list if f.endswith('.dotm')]
    searches = 0
    matches = 0

    for file_name in file_list:
        full_path = os.path.join(args.dir, file_name)
        with zipfile.ZipFile(full_path) as z:
            names = z.namelist()

            # incrementing search
            searches += 1

            with z.open('word/document.xml') as doc:
                # looping through each line
                for line in doc:
                    text_position = line.find(args.text)
                    if text_position >= 0:

                       # increment matches if found
                        matches += 1

                        print("Match found in file {}".format(file_name))
                        print("... {} ...".format(
                            line[text_position-40: text_position+41]))

    print("Files searched: {}".format(searches))
    print("Files matched: {}".format(matches))


if __name__ == '__main__':
    main(sys.argv[1:])
