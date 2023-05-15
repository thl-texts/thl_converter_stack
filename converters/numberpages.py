#!env/bin/python
"""
A converter that numbers unnumbered milestones sequentially
"""
from lxml import etree
from .baseconverter import BaseConverter
from os import path, walk, mkdir
from re import match, search
from shutil import copy


class NumberPages(BaseConverter):
    def __init__(self, args, extras=None):
        super().__init__(args, extras)
        self.stnum = args.start
        self.add_first = '-af' in self.other_settings
        self.dtd_statement = ''
        self.source_entity = ''

    def convert(self):
        # option -walk means to walk the given in directory and
        # copy the structure into the out directory with files converted
        if '-walk' in self.other_settings:
            for dirpath, dirs, files in walk(self.indir):
                dirs.sort()
                files.sort()
                for adir in dirs:
                    fulldirpath = path.join(self.outdir, adir)
                    if not path.exists(fulldirpath):
                        mkdir(fulldirpath)

                xmlfiles = [f for f in files if f.endswith('.xml')]
                for fpath in xmlfiles:
                    inpath = path.join(dirpath, fpath)
                    outpath = inpath.replace(self.indir, self.outdir)
                    self.convert_tree_doc(inpath, outpath)
            print("")

        else:
            # Otherwise convert as normal all files in workspace/in directory with converted files in ../out
            super().convert()

    def convertdoc(self):
        self.current_file_path = path.join(self.indir, self.current_file)
        self.loadxml()
        if not self.xmlroot == 'error':
            self.number_milestones()
            self.write_xml()
        else:
            copy(self.current_file_path, path.join(self.outdir, self.current_file))

    def convert_tree_doc(self, infile, outfile):
        print(f"\rConverting: {infile}      ", end="")
        self.current_file_path = infile
        self.loadxml()
        if not self.xmlroot == 'error':
            self.number_milestones()
            self.walk_write_xml(outfile)

        else:
            copy(infile, outfile)

    def loadxml(self):
        with open(self.current_file_path, "rb") as infile:
            self.dtd_statement = b''
            self.source_entity = b''
            try:
                xmltext = infile.read().decode('utf-8')
                if '!DOCTYPE' in xmltext:
                    mtch = search(r'<\![^\]]+]>', xmltext)
                    if mtch:
                        self.dtd_statement = mtch.group(0).replace("\t", "\t").replace("\n", "\n").encode('utf-8')
                    mtch = search(r'<sourceDesc n="tibbibl">[^<]+</sourceDesc>', xmltext)
                    if mtch:
                        self.source_entity = mtch.group(0).encode('utf-8')
                infile.seek(0)
                parser = etree.XMLParser(recover=True)
                tree = etree.parse(infile, parser)
                self.xmlroot = tree.getroot()
                mtch = search(r'text-p(\d+)-\d+\.xml', self.current_file)
                if mtch:
                    self.stnum = int(mtch.group(1))
                    print("start num", self.stnum)
            except UnicodeDecodeError as ude:
                print(f"\nCould not decode file: {self.current_file_path}\n")
                self.xmlroot = 'error'

    def number_milestones(self):
        root = self.xmlroot
        rnd = int(root.get('rend')) if root.get('rend') and root.get('rend').isdigit() else False
        milestones = root.xpath('//milestone')
        line_unit = 'line'
        if self.add_first:
            for ms in milestones:
                unit = ms.get('unit')
                if 'line' in unit.lower():
                    line_unit = unit
                    break

        pnm = self.stnum - 1
        if rnd:
            pnm += rnd
        lnm = 0
        for ms in milestones:
            unit = ms.get('unit')
            if 'page' in unit.lower():
                pnm += 1
                lnm = 0
                ms.set('n', str(pnm))
                if self.add_first:
                    lnms = etree.XML(f'<milestone unit="{line_unit}" n="{pnm}.1" />')
                    ms.addnext(lnms)
                    lnm = 1

            elif 'line' in unit.lower():
                lnm += 1
                ms.set('n', f"{pnm}.{lnm}")

    def write_xml(self):
        fpth = path.join(self.outdir, self.current_file)
        while path.isfile(fpth) and not self.overwrite:
            userin = input("The file {} already exists in the out folder. "
                           "Overwrite it (y/n/q): ".format(self.current_file))
            if userin == 'y':
                break
            elif userin == 'n':
                fname = input("Enter a new file name: ")
                fpth = path.join(self.outdir, fname)
            else:
                exit(0)
        with open(fpth, "wb") as outfile:
            xmlstring = etree.tostring(self.xmlroot,
                                       pretty_print=True,
                                       encoding='utf-8',
                                       xml_declaration=True)
            if len(self.dtd_statement) > 0:
                xmldecl = b'<?xml version=\'1.0\' encoding=\'utf-8\'?>\n'
                xmlstring = xmlstring.replace(xmldecl,
                                  xmldecl + self.dtd_statement.replace(b'\t', b''))
                src_ent = b'<sourceDesc n="tibbibl"/>'
                xmlstring = xmlstring.replace(src_ent, self.source_entity)
            outfile.write(xmlstring)

    def walk_write_xml(self, outfile):
        with open(outfile, "wb") as outstream:
            xmlstring = etree.tostring(self.xmlroot,
                                       pretty_print=True,
                                       encoding='utf-8',
                                       xml_declaration=True)
            if len(self.dtd_statement) > 0:
                xmldecl = b'<?xml version=\'1.0\' encoding=\'utf-8\'?>\n'
                xmlstring = xmlstring.replace(xmldecl,
                                  xmldecl + self.dtd_statement.replace(b'\t', b''))
                src_ent = b'<sourceDesc n="tibbibl"/>'
                xmlstring = xmlstring.replace(src_ent, self.source_entity)
            outstream.write(xmlstring)