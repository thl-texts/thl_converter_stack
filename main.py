#!env/bin/python

import argparse
from base import BaseConverter


def main():
    """ Parses arguments and calls convertDoc() on all documents listed """

    # Generate the arg parser and options
    parser = argparse.ArgumentParser(description='Convert THL Word marked up documents to THL TEI XML')
    parser.add_argument('-i', '--indir',
                        default='./workspace/in',
                        help='The relative path to the in-folder containing files to be converted')
    parser.add_argument('-o', '--out',
                        default='./workspace/out',
                        help='The relative path to the out-folder where converted files are written')
    parser.add_argument('-l', '--log',
                        default='./workspace/logs',
                        help='The relative path to the out-folder where converted files are written')
    parser.add_argument('-mtf', '--metafields',
                        action='store_true',
                        help='List the metadata fields in the template')
    parser.add_argument('-t', '--template',
                        default='tib_text_template.xml',
                        help='Name of template file in template folder')
    parser.add_argument('-dtd', '--dtdpath',
                        default='http://texts.thlib.org/cocoon/texts/catalogs/',
                        help='Path to the xtib3.dtd to add to the xmlfile')
    parser.add_argument('-d', '--debug',
                        action="store_true",
                        help='Whether to debug')
    args = parser.parse_args()

    converter = BaseConverter(args)
    print(converter)
    converter.convert()
    print("***********************************")


if __name__ == "__main__":
    main()

