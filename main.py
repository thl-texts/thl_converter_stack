#!env/bin/python

import argparse
from textconverter import TextConverter
from digitalpages import DigitalPages


def main():
    """ Parses arguments and calls convertDoc() on all documents listed """

    # Generate the arg parser and options
    parser = argparse.ArgumentParser(description='Convert THL Word marked up documents to THL TEI XML')
    parser.add_argument('-d', '--debug',
                        action="store_true",
                        help='Whether to debug')
    parser.add_argument('-dtd', '--dtdpath',
                        default='http://texts.thlib.org/cocoon/texts/catalogs/',
                        help='Path to the xtib3.dtd to add to the xmlfile')
    parser.add_argument('-be', '--bibl-entity',
                        action="store_true",
                        default=False,
                        help="Whether to create an entity for the bibl file base on filename")
    parser.add_argument('-e', '--edition-sigla',
                        default='',
                        help="The main edition sigla to be used for a lemma readings")
    parser.add_argument('-i', '--indir',
                        default='./workspace/in',
                        help='The relative path to the in-folder containing files to be converted. '
                             'Defaults to ./workspace/in')
    parser.add_argument('-l', '--log',
                        default='./workspace/logs',
                        help='The relative path to the out-folder where converted files are written. '
                             'Defaults to ./workspace/logs')
    parser.add_argument('-mtf', '--metafields',
                        action='store_true',
                        help='List the metadata fields in the template')
    parser.add_argument('-o', '--out',
                        default='./workspace/out',
                        help='The relative path to the out-folder where converted files are written. '
                             'Defaults to ./workspace/out')
    parser.add_argument('-opt', '--options',
                        default='',
                        help='JSON String of options for each converter')
    parser.add_argument('-ow', '--overwrite',
                        action='store_true',
                        help='Overwrite XML files by the same name in out directory')
    parser.add_argument('-t', '--template',
                        default='tib_text.xml',
                        help='Name of template file in template folder')
    # Make type the only positional that defaults to word-2-xml
    parser.add_argument('-tp', '--type',
                        default="word-2-xml",
                        help='Type of conversion to perform')

    args = parser.parse_args()

    # Initialize appropriate converter for type
    if args.type == 'digpage':
        print("Digital Page conversions!")
        converter = DigitalPages(args)
    else:
        print("Word to XML Conversion")
        converter = TextConverter(args)

    # Do the Conversion
    converter.convert()
    print("***********************************")


if __name__ == "__main__":
    main()

