#!env/bin/python
"""
A convert that just addes digital pages to word docs
"""
from baseconverter import BaseConverter
from os import path
import docx
import re


class DigitalPages(BaseConverter):
    outfile = ""

    def convertdoc(self):
        self.current_file_path = path.join(self.indir, self.current_file)
        self.worddoc = docx.Document(self.current_file_path)
        lpp = 15
        tpl = 15
        pgct = 0
        lnct = 1
        ct = 0
        tsekpattern = r'[ཀ-࿿]་|[ཀ-࿿]།|ག\s'
        for p in self.worddoc.paragraphs:
            psnm = p.style.name
            if 'Heading' not in psnm:
                # print(psnm, len(p.runs))
                for r in p.runs:
                    insertPoints = []
                    if pgct == 0:
                        pgct += 1
                        insertPoints.append({
                            'page': pgct,
                            'line': lnct,
                            'match': 0
                        })
                    tseklist = re.finditer(tsekpattern, r.text)
                    for tsek in tseklist:
                        ct += 1
                        if ct == tpl:
                            if lnct == lpp:
                                pgct += 1
                                lnct = 1
                                insertPoints.append({
                                    'page': pgct,
                                    'line': lnct,
                                    'match': tsek
                                })
                            else:
                                lnct += 1
                                insertPoints.append({
                                    'page': pgct,
                                    'line': lnct,
                                    'match': tsek
                                })
                            ct = 0
                    insertPoints.reverse()
                    rtxt = r.text
                    for ip in insertPoints:
                        stind = ip['match'] if isinstance(ip['match'], int) else ip['match'].span()[1]
                        lnnum = ip['line']
                        pgnum = ip['page']
                        ms = ''
                        if lnnum == 1:
                            ms += '[TDP {}]'.format(pgnum)
                        ms += '[TDL {}]'.format(lnnum)
                        rtxt = rtxt[:stind] + ms + rtxt[stind:]
                    r.text = rtxt
            # Need to go through document again and find all [DP 1] and [DL 13] etc. and apply character styles:
            # "Page Number" for pages and "Line Number" for lines. (Check if these are the standard style names)

        self.outfile = path.join(self.outdir, self.current_file.replace('.doc', '-out.doc'))
        self.worddoc.save(self.outfile)
