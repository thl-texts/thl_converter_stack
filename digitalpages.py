#!env/bin/python
"""
A convert that just addes digital pages to word docs
"""
from baseconverter import BaseConverter
from os import path
import docx
import re
import logging


class DigitalPages(BaseConverter):
    outfile = ""
    lines_per_page = 15
    tsek_per_line = 20
    digpage_sigla = 'TDP'
    digline_sigla = 'TDL'
    digpage_style_id = 'page number'
    digline_style_id = 'line number'
    style_id_list = [digpage_style_id, digline_style_id]

    def convertdoc(self):
        self.current_file_path = path.join(self.indir, self.current_file)
        self.worddoc = docx.Document(self.current_file_path)

        # stylist = []
        # for st in self.worddoc.styles:
        #     stylist.append("{} | {}".format(st.name, st.style_id))
        #     if st.name.lower() == 'line number' or st.name.lower() == 'LineNumber':
        #         print(st.name)
        # stylist.sort()
        # for st in stylist:
        #     print(st)
        # exit(0)

        pgct = 0                    # Page start 0
        lnct = 1                    # Line start 1
        ct = 0
        tsekpattern = r'[ཀ-࿿]་|[ཀ-࿿]།|ག\s'
        pct = 0
        totalpct = len(self.worddoc.paragraphs)
        print("Inserting milestones ...")
        for p in self.worddoc.paragraphs:
            pct += 1
            print("\rDoing paragraph {} of {}        ".format(pct, totalpct), end="")
            psnm = p.style.name
            if 'Heading' not in psnm:
                # print(psnm, len(p.runs))
                for r in p.runs:
                    insertPoints = []
                    rstyname = r.style.style_id
                    if "Default" not in rstyname:
                        logging.log(logging.DEBUG, "Style ID: {}".format(rstyname))

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
                        if ct == self.tsek_per_line:
                            if lnct == self.lines_per_page:
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
                            ms += '[{} {}]'.format(self.digpage_sigla, pgnum)
                        ms += '[{} {}.{}]'.format(self.digline_sigla, pgnum, lnnum)
                        rtxt = rtxt[:stind] + ms + rtxt[stind:]
                    r.text = rtxt

        # Need to go through document again and find all [DP 1] and [DL 13] etc. and apply character styles:
        # "Page Number" for pages and "Line Number" for lines. (Check if these are the standard style names)
        # Style ID: PageNumber and  Style ID: LineNumber
        pi = 0
        digpage_style = self.worddoc.styles[self.digpage_style_id]
        digline_style = self.worddoc.styles[self.digline_style_id]

        digmile_match = r'\[({})\s+\d+\]|\[({})\s+\d+\.\d+\]'.format(self.digpage_sigla, self.digline_sigla)
        ct = 0
        totalct = len(self.worddoc.paragraphs)
        print("\nApplying styles ...")
        for p in self.worddoc.paragraphs:
            ct += 1
            print("\rparagraph {} of {}        ".format(ct, totalct), end="")
            psnm = p.style.name
            if 'Heading' not in psnm:
                myruns = p.runs  # get all runs
                p.clear()        # remove runs to reinsert with splits
                for r in myruns:
                    rtxt = r.text
                    rsty = r.style
                    if r.style.name not in self.style_id_list:
                        mtch = re.search(digmile_match, rtxt)
                        while mtch:
                            (mst, men) = mtch.span()
                            mtxt = mtch.group()
                            sig = mtch.group(1)
                            milesty = digpage_style if sig == self.digpage_sigla else digline_style
                            beftxt = rtxt[:mst]
                            afttxt = rtxt[men:]
                            if len(beftxt) > 0:
                                p.add_run(beftxt, rsty)
                            p.add_run(mtxt, milesty)
                            rtxt = afttxt
                            mtch = re.search(digmile_match, rtxt)
                        if len(rtxt) > 0:
                            p.add_run(rtxt, rsty)
                    else:
                        p.add_run(rtxt, rsty)

        print("\n")
        self.outfile = path.join(self.outdir, self.current_file.replace('.doc', '-out.doc'))
        self.worddoc.save(self.outfile)
