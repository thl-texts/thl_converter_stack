#!env/bin/python
"""
The beginning of a base converter ancestor class copied from the TextConverter (3/118, 2021)
But not yet finalized
"""
import os
import logging
from .styleelements import fontSame

TEMPLATE_FOLDER = 'templates'


class BaseConverter:
    def __init__(self, args, other_settings=None):
        self.args = args
        self.files = []
        self.current_file = ''
        self.current_file_path = ''
        self.indir = args.indir
        self.infile_ext = args.extension
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
        self.debug_store = []
        self.worddoc = None
        self.nsmap = None
        self.metatable = None
        self.footnotes = []
        self.endnotes = []
        self.xmlroot = None
        self.headstack = []
        self.current_el = None
        self.pindex = -1
        self.edsig = ''
        self.chapnum = None
        self.textid = ''
        self.log = args.log
        self.loglevel = logging.DEBUG if self.debug else logging.WARN
        logging.basicConfig(level=self.loglevel)
        self.other_settings = other_settings

    def getfiles(self):
        files_in_dir = os.listdir(self.indir)
        files_in_dir.sort()
        for sfile in files_in_dir:
            if sfile.endswith(self.infile_ext) and not sfile.startswith('~'):
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
            if self.debug:
                self.setlog()
            self.convertdoc()

    def convertdoc(self):
        pass

    def merge_runs(self):
        '''
        Take a document and go through all runs in all paragraphs, if two consecutive runs have the same style, then merge them

        :param doc:
        :return:
        '''
        totp = len(self.worddoc.paragraphs)
        ct = 0
        for para in self.worddoc.paragraphs:
            ct += 1
            print("\rMerging runs: {}%".format(int(ct/totp*100)), end='')
            runs2remove = []
            lastrun = False
            # Merge runs with same style
            for n, r in enumerate(para.runs):
                if lastrun is False:
                    # if false no last run to compare, set lastrun
                    lastrun = r
                elif not fontSame(lastrun, r):
                    lastrun = r
                elif r.style.name == lastrun.style.name:
                    # Otherwise is charstyle and font characteristics are the same, append the two
                    lastrun.text += r.text
                    runs2remove.append(r)
                else:
                    # if style name is different and font characteristics are the same, start a new run (lastrun = r)
                    lastrun = r
            # Remove all runs thus merged
            for rr in runs2remove:
                el = rr._element
                el.getparent().remove(el)
        print("")
