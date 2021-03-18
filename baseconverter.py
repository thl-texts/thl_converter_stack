#!env/bin/python
"""
The beginning of a base converter ancestor class copied from the TextConverter (3/118, 2021)
But not yet finalized
"""
import os
import logging
import re
import zipfile
import docx
from lxml import etree

from datetime import date
from styleelements import getStyleElement

TEMPLATE_FOLDER = 'templates'


class BaseConverter:
    def __init__(self, args):
        self.files = []
        self.current_file = ''
        self.current_file_path = ''
        self.indir = args.indir
        if not os.path.isdir(self.indir):
            raise NotADirectoryError("The in path, {}, is not a directory".format(self.indir))
        self.getfiles()
        self.outdir = args.out
        self.overwrite = args.overwrite
        if not os.path.isdir(self.outdir):
            raise NotADirectoryError("The out path, {}, is not a directory".format(self.outdir))
        self.metafields = args.metafields if args.metafields else False
        self.template = args.template
        self.xmltemplate = ''
        self.dtdpath = args.dtdpath
        self.debug = args.debug if args.debug else False
        self.worddoc = None
        self.metatable = None
        self.footnotes = []
        self.endnotes = []
        self.xmlroot = None
        self.headstack = []
        self.current_el = None
        self.pindex = -1
        self.edsig = ''
        self.chapnum = None

        self.log = args.log
        self.loglevel = logging.DEBUG if self.debug else logging.WARN
        logging.basicConfig(level=self.loglevel)

    def getfiles(self):
        files_in_dir = os.listdir(self.indir)
        files_in_dir.sort()
        for sfile in files_in_dir:
            if sfile.endswith(".docx") and not sfile.startswith('~'):
                self.files.append(sfile)

    def setlog(self):
        # fname = os.path.split(fname)[1].replace('.docx', '') + '.log'
        logpath = os.path.join(self.log, self.current_file.replace('docx', 'log'))
        if self.debug:
            print("Log file for {} is: {}".format(self.current_file, logpath))
        loghandler = logging.FileHandler(logpath, 'w')
        log = logging.getLogger()
        for hdlr in log.handlers[:]:
            log.removeHandler(hdlr)
        log.addHandler(loghandler)

    def convert(self):
        for fl in self.files:
            print("\n======================================\nConverting file: {}".format(fl))
            self.current_file = fl
            self.setlog()
            self.convertdoc()

    def convertdoc(self):
        pass
