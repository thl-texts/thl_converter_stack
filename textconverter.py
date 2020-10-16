#!env/bin/python
"""
The main Word to Text converter for THL-specific TEI texts. Could be inherited by other types of converters
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


class TextConverter:

    def __init__(self, args):
        self.files = []
        self.current_file = ''
        self.current_file_path = ''
        self.indir = args.indir
        if not os.path.isdir(self.indir):
            raise NotADirectoryError("The in path, {}, is not a directory".format(self.indir))
        self.getfiles()
        self.outdir = args.out
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

        self.log = args.log
        self.loglevel = logging.DEBUG if self.debug else logging.WARN
        logging.basicConfig(level=self.loglevel)

    def getfiles(self):
        for sfile in os.listdir(self.indir):
            if sfile.endswith(".docx") and not sfile.startswith('~'):
                self.files.append(sfile)

    def setlog(self):
        # fname = os.path.split(fname)[1].replace('.docx', '') + '.log'
        logpath = os.path.join(self.log, self.current_file.replace('docx', 'log'))
        print(logpath)
        if self.debug:
            print("Log file for {} is: {}".format(self.current_file, logpath))
        loghandler = logging.FileHandler(logpath, 'w')
        log = logging.getLogger()
        for hdlr in log.handlers[:]:
            log.removeHandler(hdlr)
        log.addHandler(loghandler)

    def convert(self):
        for fl in self.files:
            print("Converting file: {}".format(fl))
            self.current_file = fl
            self.setlog()
            self.convertdoc()
            self.writexml()

    def convertdoc(self):
        self.current_file_path = os.path.join(self.indir, self.current_file)
        self.worddoc = docx.Document(self.current_file_path)
        self.merge_runs()
        self.pre_process_notes()
        self.createxml()

        # TODO: add more conversion code here. Iterate through paras then runs. Using stack method
        totalp = len(self.worddoc.paragraphs)
        ct = 0
        for index, p in enumerate(self.worddoc.paragraphs):
            ct += 1
            print("\rDoing paragraph {} of {}  ".format(ct, totalp), end="")
            self.pindex = index
            if isinstance(p, docx.text.paragraph.Paragraph):
                self.convertpara(p)
            else:
                self.mywarning("Warning: paragraph ({}) is not a docx paragraph cannot convert".format(p))
        print("")

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
                if lastrun is not False and r.style.name == lastrun.style.name:
                    lastrun.text += r.text
                    runs2remove.append(r)
                else:
                    lastrun = r
            # Remove all runs thus merged
            for rr in runs2remove:
                el = rr._element
                el.getparent().remove(el)
        print("")

    def pre_process_notes(self):
        """
        Preprocess footnotes and endnotes

        TODO: Check how footnote XML deals with italics, bold, etc and convert to TEI compliant

        :return:
        """
        zipdoc = zipfile.ZipFile(self.current_file_path)

        # write content of endnotes.xml into self.footnotes[]
        fnotestxt = zipdoc.read('word/footnotes.xml')
        xml_fn_root = etree.fromstring(fnotestxt)
        # To output the footnote XML file from Word uncomment the lines below:
        # with open('./workspace/logs/footnotes-test.xml', 'wb') as xfout:
        #     xfout.write(etree.tostring(xml_fn_root))
        fnindex = 0
        fnotes = xml_fn_root.findall('w:footnote', xml_fn_root.nsmap)
        for f in fnotes:
            if fnindex > 1:  # The first two "footnotes" are the separation and continuation lines
                text = f.findall('.//w:t', xml_fn_root.nsmap)
                s = ""
                wdschema = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                for t in text:
                    ttxt = t.text
                    prev = t.getprevious()  # This returns <w:rPr> or None
                    if prev is not None:
                        tsty = []
                        for pc in prev.getchildren():
                            pcstyle = re.sub(r'\{[^\}]+\}', '', pc.tag)
                            if pcstyle == 'rStyle':
                                if pc.get(wdschema + 'val') == 'X-EmphasisStrong':
                                    tsty.append('strong')
                            if pcstyle == 'i':
                                tsty.append('weak')
                            if pcstyle == 'u':
                                tsty.append('underline')
                        ttxt = '<hi rend="{}">{}</hi>'.format(' '.join(tsty), ttxt)
                    s += ttxt
                note_el = etree.XML('<note type="footnote">{}</note>'.format(s))
                self.footnotes.append(note_el)
            fnindex += 1

        # write content of endnotes.xml into self.endnotes[]
        xml_content = zipdoc.read('word/endnotes.xml')
        xml_en_root = etree.fromstring(xml_content)
        enindex = 0
        enotes = xml_en_root.findall('w:endnote', xml_en_root.nsmap)
        for f in enotes:
            if enindex > 1:
                text = f.findall('.//w:t', xml_en_root.nsmap)
                s = ""
                wdschema = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                for t in text:
                    ttxt = t.text
                    prev = t.getprevious()  # This returns <w:rPr> or None
                    if prev is not None:
                        tsty = []
                        for pc in prev.getchildren():
                            pcstyle = re.sub(r'\{[^\}]+\}', '', pc.tag)
                            if pcstyle == 'rStyle':
                                if pc.get(wdschema + 'val') == 'X-EmphasisStrong':
                                    tsty.append('strong')
                            if pcstyle == 'i':
                                tsty.append('weak')
                            if pcstyle == 'u':
                                tsty.append('underline')
                        ttxt = '<hi rend="{}">{}</hi>'.format(' '.join(tsty), ttxt)
                    s += ttxt
                note_el = etree.XML('<note type="endnote">{}</note>'.format(s))
                self.endnotes.append(note_el)
            enindex += 1
        zipdoc.close()

    def createxml(self):
        with open(os.path.join(TEMPLATE_FOLDER, self.template), 'r') as tempstream:
            self.xmltemplate = tempstream.read()
            self.metatable = self.worddoc.tables[0] if len(self.worddoc.tables) else False
            if self.metatable:
                self.createmeta()  # Separated out to be overriden. Creates the XML string with metadata inserted
            # Take self.xmltemplate the xml string and convert to an XML object
            xmldoc = bytes(bytearray(self.xmltemplate, encoding='utf-8'))
            # create lxml element tree from metadata info
            parser = etree.XMLParser(ns_clean=True, recover=True, encoding='utf-8')
            self.xmlroot = etree.fromstring(xmldoc, parser)

    def createmeta(self):
        """
        This table creates the xmlstring by filling in the template string with the table values based on the
        table labels. It uses the classes self.metatable property which is the Word docs metatable and
        self.template which is the string read in from the XML template file indicated in the settings

        NOTE: This function should be overridden by particular template converters to customize
        TODO: Create a generalized version to put here and use this for the chris_old template
        :return:
        """
        wordtable = self.metatable
        # Fill out metadata matching on string in wordtable with {strings} in template (teiHeader.dat)
        xmltext = self.xmltemplate.replace("{Digital Creation Date}", str(date.today()))
        problems_on = False
        tablerows = len(wordtable.rows)
        for rwnum in range(0, tablerows):
            try:
                if wordtable._column_count == 2:
                    label = wordtable.cell(rwnum, 0).text.strip()
                    rowval = wordtable.cell(rwnum, 1).text.strip()
                elif wordtable._column_count == 3:
                    label = wordtable.cell(rwnum, 0).text.strip()
                    rowval = wordtable.cell(rwnum, 1).text.strip()
                    msg = "***** NOTICE: Need to update code for processing 3 column metadata tables!!!!! *******"
                    self.mywarning(msg)
                else:
                    raise (ValueError('Metadata table does not have the right number of columns. Must be 2 or 3'))

                # All Uppercase are Headers in the table skip
                if label.isupper():
                    if label == 'PROBLEMS':
                        problems_on = True
                    continue  # Skip labels
                if problems_on:
                    if '<encodingDesc>' not in xmltext:
                        xmltext = xmltext.replace('</fileDesc>',
                                                    '</fileDesc><encodingDesc><editorialDecl ' +
                                                    'n="problems"><interpretation n="{}">'.format(label) +
                                                    '<p>{}</p>'.format(rowval) +
                                                    '</interpretation></editorialDecl></encodingDesc>')
                    else:
                        xmltext = xmltext.replace('</interpretation></editorialDecl>',
                                                    '</interpretation><interpretation n="{}">'.format(label) +
                                                    '<p>{}</p>'.format(rowval) +
                                                    '</interpretation></editorialDecl>')
                temppt = label.split(' (')  # some rows have " (if applicable)" or possible some other instruction
                label = temppt[0].replace('\xa0', ' ').strip()
                label = label.replace(' (if applicable)', '')
                label = label.replace('Call་number', 'Call-number')
                srclbl = "{" + label + "}"
                xmltext = xmltext.replace(srclbl, rowval)

            except IndexError as e:
                logging.error("Index error in iterating wordtable: {}".format(e))
            except TypeError as e:
                logging.error("Type error in iterating wordtable: {}".format(e))

        self.xmltemplate = re.sub(r'{([^}]+)}', r'<!--\1-->', xmltext)

    def convertpara(self, p):
        style_name = p.style.name
        headmtch = re.match(r'^Heading (?:Tibetan\s*)?(\d+)[\,\s]*(Front|Body|Back)?', style_name)
        if headmtch:
            self.do_header(p, headmtch)

        elif "List" in style_name:
            self.do_list(p)

        elif "Verse" in style_name:
            self.do_verse(p)

        elif "Citation" in style_name:
            self.do_citation(p)

        elif "Section" in style_name:
            self.do_section(p)

        elif "Speech" in style_name:
            self.do_speech(p)

        else:
            if not self.is_reg_p(style_name):
                msg = "Style {} defaulting to paragraph".format(style_name)
                self.mywarning(msg)
            self.reset_current_el()
            self.do_paragraph(p)

        # Once paragraphs have been processed. self.current_el is the element where the runs of the paragraph go
        self.iterate_runs(p)  # so all we need to do is send the word paragraph object

    def do_header(self, p, headmtch):
        """
        Method to convert heading styles into structural elements either front, body, bock or divs
        Called from convertpara() method above

        :param p: the word docx paragraph object
        :param headmtch: the re.match object group(1) is heading level, group(2) is front, body, back if it is one of those
        :return: none
        """
        hlevel = int(headmtch.group(1))
        # TODO: need to parse the p element below in case there is internal markup to put in head
        #  (i.e. don't just use p.text but it may have children)
        htext = p.text
        mtch = re.match(r'^((\d+\.?)+)', htext)
        if mtch:
            hnum = mtch.group(1)
            htext = htext.replace(hnum, '<num>{}</num>'.format(hnum))
        style_name = p.style.name
        if hlevel == 0:
            # If level is 0, its front body or back, create element and clear head stack
            fbbel = etree.XML('<{0}><head>{1}</head></{0}>'.format(headmtch.group(2).lower(), htext))
            self.xmlroot.find('text').append(fbbel)
            self.current_el = fbbel.find('head')
            self.headstack = [fbbel]
        else:
            # Otherwise we are already in front, body, or back, so create div
            currlevel = len(self.headstack) - 1  # subtract one bec. div 0 is at top of stack
            # TODO: need to parse the p element in case there is internal markup to put in head
            hdiv = etree.XML('<div n="{}"><head>{}</head><p></p></div>'.format(hlevel, htext))
            # if it's the next level deeper
            if hlevel > currlevel:
                # if new level is higher than the current level just add it to current
                if hlevel - currlevel > 1:
                    self.mywarning("Warning: Heading level skipped for {}".format(style_name, htext))
                self.headstack[-1].append(hdiv)  # append the hdiv to the previous one
                self.headstack.append(hdiv)      # add the hdiv to the stack
            # if it's the same level as current
            elif hlevel == currlevel:
                self.headstack[-1].addnext(hdiv)
                self.headstack[-1] = hdiv
            # Otherwise it's a higher level
            else:
                # because front, body, back is 0th element use hlevel to splice array
                self.headstack = self.headstack[0:hlevel]  # splice arra to parent elements
                self.headstack[-1].append(hdiv)
                self.headstack.append(hdiv)
            self.current_el = hdiv.find('head')

    def do_list(self, p):
        # Get current and previous list styles and numbers
        my_style = p.style.name
        my_num = self.getnumber(my_style) or 1
        prev_style = self.get_previous_p(True)
        prev_num = self.getnumber(prev_style)
        if not prev_num:
            if my_style == prev_style:
                prev_num = my_num
            elif "List" in prev_style:
                prev_num = 1

        # Determine the type of list and set attribute values
        rend = "bullet" if "Bullet" in my_style else "1"
        nval = ' n="1"' if rend == "1" else False
        # ptxt = p.text

        # Create element object templates (used below in different contexts). Not all are used in all cases
        itemlistel = etree.XML('<item><list rend="{}"{}><item></item></list></item>'.format(rend, nval))
        listel = etree.XML('<list rend="{}"{}><item></item></list>'.format(rend, nval))
        itemel = etree.XML('<item></item>')

        # Different Alternatives
        if not prev_num:  # new non-embeded list
            if my_num and int(my_num) > 1:
                self.mywarning("List level {} added when not in list".format(my_num))
            # Add list after current element and make it current
            self.current_el.addnext(listel)
            self.current_el = listel.find('item')

        else:   # For Lists embeded in Lists
            my_num = int(my_num)
            prev_num = int(prev_num)
            if my_num > prev_num:  # Adding/Embedding a new level of list
                if my_num > prev_num + 1:
                    self.mywarning("Skipping list Level: inserting {} at level {}".format(my_num, prev_num))
                # Append the itemlist el to the current list
                self.current_el.addnext(itemlistel)
                self.current_el = itemlistel.find('list/item')

            elif my_num == prev_num: # On same level lists
                self.current_el.addnext(itemel)
                self.current_el = itemel

            elif self.current_el is not None:  # my_num < prev_num  Returning to an ancestor list with lower list number
                lvl = prev_num
                listel = self.current_el.getparent()
                try:
                    while lvl != my_num:
                        listel = listel.getparent()
                        if listel.tag == 'list':
                            lvl -= 1
                except AttributeError:
                    self.mywarning("No iterancestors() method for current list element (text: {}). " +
                                   "Current Element Class: {}".format(p.text, self.current_el.__class__.__name__))
                listel.append(itemel)
                self.current_el = itemel
            else:  # No current el so use the last div
                self.headstack[-1].append(listel)
                self.current_el = listel.find('item')

    def do_verse(self, p):
        my_style = p.style.name
        prev_style = self.get_previous_p(True)
        vrs_el = etree.Element('p')
        self.current_el.addnext(vrs_el)
        self.current_el = vrs_el

    def do_citation(self, p):
        self.do_verse(p)

    def do_section(self, p):
        self.do_verse(p)

    def do_speech(self, p):
        self.do_verse(p)

    def do_paragraph(self, p):
        if p.text.strip() == '':
            return
        my_style = p.style.name
        prev_style = self.get_previous_p(True)
        p_el = etree.Element('p')
        if self.current_el.tag == 'div':
            self.current_el.append(p_el)
        else:
            self.current_el.addnext(p_el)
        self.current_el = p_el

    def iterate_runs(self, p):
        '''
        In old converter this was iterateRange (the interateRuns function was not called)
        :param p:
        :return:
        '''
        last_run_style = ''
        temp_el = etree.Element('temp')  # temp element to put xml element objects in
        elem = None
        if len(p.runs) == 0:
            temp_el.text = " "
            return

        if temp_el.text is None:
            temp_el.text = ""

        for run in p.runs:
            if run is None:
                continue
            rtxt = run.text
            if elem is not None and elem.text is None:
                elem.text = ""
            if elem is not None and elem.tail is None:
                elem.tail = ""

            if "Heading" in p.style.name:
                rtxt = re.sub(r'^[\d\s\.]+', '', rtxt)

            char_style = run.style.name
            is_new_style = True if char_style != last_run_style else False
            # Default Paragraph Font
            if not char_style or char_style == "" or "Default Paragraph Font" in char_style:
                if elem is None:
                    temp_el.text += rtxt
                else:
                    elem.tail += rtxt
            # Footnotes
            # elif "footnote" in char_style.lower() or "endnote" in char_style.lower():
            #     # print("Doing note Style: {}".format(char_style))
            #     note_el = self.do_footendnotes(run, elem, temp_el)
            #     if note_el:
            #         elem = note_el

            # Milestones
            elif "Page Number" in char_style or "Line Number" in char_style:
                # If there are multiple ms of the same style, they get merged in merge_runs. So split them up
                msitems = rtxt.split('][')
                for mstxt in msitems:
                    elem = self.createmilestone(char_style, mstxt)
                    temp_el.append(elem)

            elif is_new_style or elem is None or elem is False:
                new_el = getStyleElement(char_style)
                if new_el is None:
                    new_el = etree.Element('s')
                    new_comment = etree.Comment("No style definition found for style name: {}".format(char_style))
                    new_el.append(new_comment)
                logging.debug('Style element {} => {}'.format(char_style, new_el.tag))
                new_el.text = rtxt
                temp_el.append(new_el)
                elem = new_el

            else:
                elem.text += rtxt

            last_run_style = char_style
        # End of interating runs

        # Copy temp_el contents to current_el depending on whether it has children or not
        curr_child = self.current_el.getchildren() or []
        if len(curr_child) > 0:
            curr_child[-1].tail = temp_el.text
        else:
            self.current_el.text = temp_el.text
        for tempchild in temp_el.getchildren():
            self.current_el.append(tempchild)

    def do_footendnotes(self, run, elem, para_el):
        # self.mywarning("TODO: Deal with critical edition notes!!!!")
        if "endnote" in run.style.name.lower():
            note = self.endnotes.pop(0)
        else:
            note = self.footnotes.pop(0)

        if elem is not None:
            elem.append(note)
        else:
            para_el.append(note)

    @staticmethod
    def createmilestone(char_style, mstxt):
        mstype = 'line' if 'line' in char_style.lower() else 'page'  # defaults to page
        msnum = mstxt.replace('[', '').replace(']', '')   # Default backup num if regex doesn't match
        mtch = re.match(r'\[?(Page|Line)\s+([^\]]+)\]?', mstxt, re.IGNORECASE)
        if mtch:
            mstype = mtch.group(1).replace('[', '').replace(']', '')
            msnum = mtch.group(2).replace('[', '').replace(']', '')
        else:
            logging.warning("No match for milestone parts in {}".format(mstxt))
        msel = getStyleElement(char_style)
        msel.set('unit', mstype)
        sep = '.' if '.' in msnum else '-'  # Do we need to check for more separators
        pts = msnum.split(sep)
        if len(pts) > 1:
            msel.set('ed', pts[0])
            msel.set('n', pts[1])
        else:
            msel.set('n', pts[0])
        return msel

    def reset_current_el(self):
        """
        Resets the current element to the last element within the current div.
        This is used to get out of nested lists and verses

        :return:
        """
        if len(self.headstack) > 0:
            hdr = self.headstack[-1]
            children = hdr.getchildren()
            if len(children) == 0:
                self.current_el = hdr
            else:
                self.current_el = children[-1]


    def writexml(self):
        # Determine Name for Resulting XML file
        fname = self.current_file.replace('.docx', '.xml')
        fpth = os.path.join(self.outdir, fname)
        while os.path.isfile(fpth):
            userin = input("The file {} already exists. Overwrite it (y/n/q): ".format(fname))
            if userin == 'y':
                break
            elif userin == 'n':
                fname = input("Enter a new file name: ")
                fpth = os.path.join(self.outdir, fname)
            else:
                exit(0)

        # Write XML File
        with open(fpth, "wb") as outfile:
            docType = "<!DOCTYPE TEI.2 SYSTEM \"{}xtib3.dtd\">".format(self.dtdpath)
            xmlstring = etree.tostring(self.xmlroot,
                                        pretty_print=True,
                                        encoding='utf-8',
                                        xml_declaration=True,
                                        doctype=docType)
            outfile.write(xmlstring)

    #  HELPER METHODS
    def get_previous_p(self, as_style=False):
        pind = self.pindex - 1
        if pind > -1:
            prevp = self.worddoc.paragraphs[pind]
            if as_style:
                return prevp.style.name
            else:
                return prevp
        return False

    # STATIC HELPER METHODS
    @staticmethod
    def mywarning(msg):
        print(msg)
        logging.warning(msg)

    @staticmethod
    def getnumber(stynm):
        mtch = re.match(r'[^\d]+\s+(\d+)', stynm)
        if mtch:
            return mtch.group(1)
        return False

    @staticmethod
    def is_reg_p(style_name):
        # TODO: Check if I need to account for Paragraph Citation. Is that a "regular paragragh"?
        if "Paragraph" not in style_name and "Outline" not in style_name and "Normal" not in style_name:
            return False
        return True