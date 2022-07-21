#!env/bin/python
"""
The main Word to Text converter for THL-specific TEI texts. Could be inherited by other types of converters
"""
import os
import logging
import re
import unicodedata
import zipfile
import docx
from lxml import etree

from datetime import date
from styleelements import getStyleElement, fontSame, getFontElement

TEMPLATE_FOLDER = 'templates'
IGNORABLE_STYLES = ['Paragraph Char', 'List Bullet Char']


def get_lang_by_char(chr):
    unm = unicodedata.name(chr)
    if "TIBETAN" in unm:
        return "tib"
    if "CHINESE" in unm:
        return "chi"
    if "DEVANAGARI" in unm:
        return "san"
    return ''


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
        self.textid = ''

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
            self.assignids()
            self.tidyxml()
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
        in_app = False
        app_ps = []
        for index, p in enumerate(self.worddoc.paragraphs):
            ct += 1
            print("\rDoing paragraph {} of {}  ".format(ct, totalp), end="")
            self.pindex = index
            if isinstance(p, docx.text.paragraph.Paragraph):
                ptxt = p.text
                if len(ptxt) > 0 and ptxt[0] == '{' and '}' not in ptxt:
                    logging.info('Multiline apparatus begins: ' + ptxt)
                    in_app = True
                if in_app:
                    app_ps.append(p)
                    if ptxt[0] == '}':
                        self.process_multiline_app(app_ps)
                        logging.info("Multiline apparatus finished: " + ptxt)
                        in_app = False
                        apps_ps = []
                else:
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

    def pre_process_notes(self):
        """
        Preprocess footnotes and endnotes

        TODO: Check how footnote XML deals with italics, bold, etc and convert to TEI compliant

        :return:
        """
        zipdoc = zipfile.ZipFile(self.current_file_path)

        # write content of endnotes.xml into self.footnotes[]
        fntfile = 'word/footnotes.xml'
        if fntfile in zipdoc.namelist():
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
                    plains = ""
                    wdschema = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                    for t in text:
                        ttxt = t.text
                        plains += ttxt
                        prev = t.getprevious()  # This returns <w:rPr> or None
                        if prev is not None:
                            tsty = []
                            lang = ""
                            for pc in prev.getchildren():
                                pcstyle = re.sub(r'\{[^\}]+\}', '', pc.tag)
                                if pcstyle == 'rStyle':
                                    if pc.get(wdschema + 'val') == 'X-EmphasisStrong':
                                        tsty.append('strong')
                                elif pcstyle == 'i':
                                    tsty.append('weak')
                                elif pcstyle == 'u':
                                    tsty.append('underline')
                                elif pcstyle == 'lang':
                                    if pc.get(wdschema + 'bidi') == 'bo-CN':
                                        lang = ' lang="tib"'
                            attr = "" if len(tsty) == 0 else ' rend="{}"'.format(' '.join(tsty))
                            if lang == "":
                                langcode = get_lang_by_char(ttxt[0])
                                if len(langcode) > 0:
                                    lang = ' lang="{}"'.format(langcode)
                            attr += lang
                            ttxt = '<hi{}>{}</hi>'.format(attr, ttxt) if len(attr) > 0 else ttxt
                        s += ttxt
                        # TODO: add regex here to find <hi> with the same rend next to each other and merge them,
                        #  e.g. <hi lang="tib">ཡོད</hi><hi lang="tib">།</hi>
                    note_el = etree.XML('<note type="footnote">{}<rs>{}</rs></note>'.format(s, plains))
                    self.footnotes.append(note_el)
                fnindex += 1

        endntfile = 'word/endnotes.xml'
        if endntfile in zipdoc.namelist():
            # write content of endnotes.xml into self.endnotes[]
            xml_content = zipdoc.read(endntfile)
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
        template_path = os.path.join(TEMPLATE_FOLDER, self.template)
        with open(template_path, 'r') as tempstream:
            if self.debug:
                print(f"Template file: {template_path}")
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

        NOTE: This function can be overridden by particular template converters to customize

        :return:
        """
        wordtable = self.metatable
        # Fill out metadata matching on string in wordtable with {strings} in template (teiHeader.dat)
        xmltext = self.xmltemplate.replace("{Digital Creation Date}", str(date.today()))
        problems_on = False
        tablerows = len(wordtable.rows)
        problems = []
        print("Process Metadata Table ... {}".format(tablerows))
        for rwnum in range(0, tablerows):
            try:
                if wordtable._column_count < 2:
                    label = wordtable.cell(rwnum, 0).text.strip()
                    logging.warning("Row {} of metadata table has too few cells".format(rwnum))
                    continue

                elif wordtable._column_count == 2:
                    label = wordtable.cell(rwnum, 0).text.strip()
                    rowval = wordtable.cell(rwnum, 1).text.strip()

                elif wordtable._column_count == 3:
                    label = wordtable.cell(rwnum, 0).text.strip()
                    rowval = wordtable.cell(rwnum, 1).text.strip()
                    self.mywarning("Notice: Not handling third column of row {}!".format(rwnum))

                elif wordtable._column_count == 4 and wordtable.cell(rwnum, 0).text == wordtable.cell(rwnum, 2).text:
                    # Not giving warning because this seems to be a Word or pydoc weirdness
                    # In the stylized metadata table, it detects four columns per row
                    # with the first 3 all containing the same label. So just using last two.
                    label = wordtable.cell(rwnum, 2).text.strip()
                    rowval = wordtable.cell(rwnum, 3).text.strip()

                else:
                    rowcells = wordtable.rows[rwnum].cells
                    label = wordtable.cell(rwnum, 0).text.strip() if len(rowcells) > 0 else ""
                    rowval = wordtable.cell(rwnum, 1).text.strip() if len(rowcells) > 1 else label
                    # self.mywarning("Row {} of metadata table has too many ({}) cells. Using first two".format(rwnum))

                if label == "Text ID":
                    self.textid = rowval

                # All Uppercase are Headers in the table skip
                if label.isupper():
                    if label == 'PROBLEMS':
                        problems_on = True
                    continue  # Skip labels
                if problems_on:
                    print("Processing Problems: ", end="")
                    for cellp in wordtable.cell(rwnum, 0).paragraphs:
                        problems.append(cellp.text.strip())
                    if rwnum < tablerows - 1:
                        continue  # iterate through all the problems which should always be last in the table
                    print(" Processed {} problems".format(len(problems)))
                    label = "Problems"   # the string {Problems} should be in template where problems go
                    problems = [f"<p>{p}</p>" for p in problems]
                    rowval = "\n".join(problems)

                temppt = label.split(' (')  # some rows have " (if applicable)" or possible some other instruction

                # Normalizing Label
                label = temppt[0].replace('\xa0', ' ').strip()
                label = label.replace(' (if applicable)', '')
                label = label.replace('Call་number', 'Call-number')
                # 'Cover Title Tib' is in XML comment while 'Title on Cover' is in form (maybe a later correction)
                if label in ('Title on Cover', 'Title on Cover Page', 'Title on Title Page'):
                    label = 'Cover Title Tib'
                # Normalize 'Cover Page Title' and 'Title Page Title' to 'Cover Title' in all occurrences
                if label != 'Cover Page':
                    label = label.replace('Cover Page', 'Cover').replace('Title Page', 'Cover')
                srclbl = "{" + label + "}"

                # Replace occurence of label in XML header
                xmltext = xmltext.replace(srclbl, rowval)

                # Deal with edition sigla
                if label.lower() == 'edition sigla':
                    self.edsig = rowval

                # Deal with Chapter Number
                if label.lower() == 'chapter number':
                    self.chapnum = rowval

            except IndexError as e:
                logging.error("Index error in iterating wordtable: {}".format(e))
            except TypeError as e:
                logging.error("Type error in iterating wordtable: {}".format(e))

        # if there are no problems the {problems} needs to be replaced with an empty paragraph.
        xmltext = xmltext.replace('{Problems}', '<p>No problems</p>')
        self.xmltemplate = re.sub(r'{([^}]+)}', r'<!--\1-->', xmltext)

    def convertpara(self, p):
        style_name = p.style.name
        headmtch = re.match(r'^Heading (?:Tibetan\s*)?(\d+)[\,\s]*(Front|Body|Back)?', style_name)
        if headmtch:
            self.do_header(p, headmtch)

        elif len(self.headstack) == 0:
            # if there is not yet a headstack then it's notes at beginning of document that should be ignored
            ptxt = p.text[0:150] if len(p.text) > 150 else p.text
            print(f"Skipping Paragraph at beginning: {ptxt}")
            return

        elif "List" in style_name:
            self.do_list(p)

        elif "Verse" in style_name:
            self.do_verse(p)  # Note this does verse citation and verse speech as well

        elif "Citation" in style_name:
            self.do_citation(p)  # Note verse citation is done above in do_verse()

        elif "Section" in style_name:
            doruns = self.do_section(p)
            if not doruns:
                return

        elif "Speech" in style_name:
            self.do_speech(p)  # verse speech is done in do_verse()

        else:
            if not self.is_reg_p(style_name):
                msg = "\n\tStyle {} defaulting to paragraph".format(style_name)
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
        style_name = p.style.name
        if hlevel == 0:
            # If level is 0, its front body or back, create element and clear head stack
            fbbel = etree.XML('<{0}><head></head></{0}>'.format(headmtch.group(2).lower())) # the match is e.g. "front"
            self.xmlroot.find('text').append(fbbel)
            self.current_el = fbbel.find('head')
            self.headstack = [fbbel]
        else:
            # Otherwise we are already in front, body, or back, so create div
            currlevel = len(self.headstack) - 1  # subtract one bec. div 0 is at top of stack
            # TODO: Need to assign ids to divs "a1", "b4", "c2" etc.
            # TODO: need to parse the p element in case there is internal markup to put in head
            hdiv = etree.XML('<div n="{}"><head></head></div>'.format(hlevel))
            # if it's the next level deeper
            if hlevel > currlevel:
                # if new level is higher than the current level just add it to current
                if hlevel - currlevel > 1:
                    self.mywarning("Warning: Heading level skipped for {}".format(style_name, p.text))
                self.headstack[-1].append(hdiv)  # append the hdiv to the previous one
                self.headstack.append(hdiv)      # add the hdiv to the stack
            # if it's the same level as current
            elif hlevel == currlevel:
                if len(self.headstack) > 0 :
                    self.headstack[-1].addnext(hdiv)
                    self.headstack[-1] = hdiv
                else:
                    errmsg = "Headstack is empty when adding div ({})\n".format(hdiv.text)
                    errmsg += "Make sure Body Heading0 is present."
                    raise ConversionException(errmsg)

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
        prev_num = self.getnumber(prev_style) if "List" in prev_style else False
        if not prev_num:
            if my_style == prev_style:
                prev_num = my_num
            elif "List" in prev_style:
                prev_num = 1

        # Determine the type of list and set attribute values
        rend = "bullet" if "Bullet" in my_style else "1"
        nval = ' n="1"' if rend == "1" else ""
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
        prev_style = self.get_previous_p(True)  # TODO: Check if current style is same (except number) with previous
        is_cite = True if 'citation' in my_style.lower() else False
        is_speech = True if 'speech' in my_style.lower() else False
        is_nested = True if 'nested' in my_style.lower() else False
        is_same = True if my_style.lower().replace('2', '1') == prev_style.lower().replace('2', '1') else False
        level = 2 if '2' in my_style else 1
        if level == 2:
            myel = etree.Element('l')
            if not is_nested and "nested" in prev_style.lower():
                self.current_el.getparent().addnext(myel)
            else:
                self.current_el.addnext(myel)
            self.current_el = myel

        elif is_cite:
            if is_same:
                markup = etree.XML('<lg><l></l></lg>')
                self.current_el.getparent().addnext(markup)
                self.current_el = markup.find('l')

            elif is_nested and "citation" in prev_style.lower() and "paragraph" in prev_style.lower():
                # when nested in paragraph citation
                markup = etree.XML('<lg><l></l></lg>')
                self.current_el.addnext(markup)
                self.current_el = markup.find('l')

            else:
                markup = etree.XML('<quote><lg><l></l></lg></quote>')
                if is_nested:
                    self.current_el.append(markup)
                else:
                    self.headstack[-1].append(markup)
                self.current_el = markup.find('lg').find('l')

        elif is_nested:
            markup = etree.XML('<lg><l></l></lg>')
            self.current_el.addnext(markup)  # if nested, current el is "l", add "lg" next to this.
            self.current_el = markup.find('l')

        else:
            if is_speech:
                markup = etree.XML('<q><lg><l></l></lg></q>')
                self.current_el = markup.find('lg').find('l')
            else:
                markup = etree.XML('<lg><l></l></lg>')
                self.current_el = markup.find('l')
            self.headstack[-1].append(markup)

    def do_citation(self, p):
        my_style = p.style.name
        prev_style = self.get_previous_p(True)  # TODO: Check if current style is same (except number) with previous
        nested = True if "nested" in my_style.lower() else False
        continued = True if "continued" in my_style.lower() else False

        if continued or (nested and my_style == prev_style):
            cite_el = etree.XML('<p rend="cont"></p>')
            if "verse" in prev_style.lower():
                self.current_el.getparent().addnext(cite_el)
            else:
                self.current_el.addnext(cite_el)
            self.current_el = cite_el
        else:
            cite_el = etree.XML('<quote><p></p></quote>')
            if nested:
                self.current_el.append(cite_el)
            elif "nested" in prev_style:
                self.current_el.getparent().addnext(cite_el)
            else:
                self.current_el.addnext(cite_el)
            self.current_el = cite_el.find('p')

    def do_section(self, p):
        my_style = p.style.name
        ptext = p.text
        nmtch = re.match(r'section\s+(\d)', my_style, re.IGNORECASE)
        if nmtch:
            n = nmtch.group(1)
            sect_el = etree.XML('<milestone unit="section" n="{}" rend="{}" />'.format(n, ptext))
            self.headstack[-1].append(sect_el)
            self.current_el = sect_el
            return False

        elif 'chapter element' in my_style.lower():
            sect_el = etree.XML('<milestone unit="section" n="cle" rend="{}" />'.format(ptext))
            self.headstack[-1].append(sect_el)
            self.current_el = sect_el
            return False

        elif 'interstitial' in my_style.lower():
            sect_el = etree.XML('<div type="interstitial"><head></head></div>')
            self.headstack[-1].append(sect_el)
            self.current_el = sect_el.find('head')
            return True

        else:
            logging.warning("Unknown section type with header: {}".format(ptext))
            # TODO: Should this be a milestone instead of a div???
            sect_el = etree.XML('<div type="section"><head></head></div>')
            self.headstack[-1].append(sect_el)
            self.current_el = sect_el.find('head')
            return True

    def do_speech(self, p):
        #  Note this is speech not already covered in verse or citation. See convertpara() method above
        my_style = p.style.name.lower()
        prev_style = self.get_previous_p(True)
        iscont = True if "continued" in my_style else False
        isnested = True if "nested" in my_style else False
        speech_el = etree.XML('<q><p></p></q>')  # use for new nested or new not nested
        if iscont:
            # iscontinued whether nested or not
            speech_el = etree.XML('<p rend="cont"></p>')
            if "nested" in prev_style and "nested" not in my_style:
                self.current_el.getparent().addnext(speech_el)
            else:
                self.current_el.addnext(speech_el)
            self.current_el = speech_el

        elif isnested:
            # initial nested speech, then the current el is a <p> within a <q> and another <q> after it
            self.current_el.addnext(speech_el)
            self.current_el = speech_el.find('p')

        else:
            # Regular "Speech Paragraph" style
            # Get out from within verse
            if self.current_el.tag == 'l':
                lgs = [lg for lg in self.current_el.iterancestors('lg')]
                self.current_el = lgs[-1] if len(lgs) > 0 else self.current_el.getparent()
            # Get out from within list
            if self.current_el.tag == 'item':
                lists = [lst for lst in self.current_el.iterancestors('list')]
                self.current_el = lists[-1] if len(lists) > 0 else self.current_el.getparent()
            # Add <q><p>... as sibling of current element.
            self.current_el.addnext(speech_el)
            self.current_el = speech_el.find('p')

    def do_paragraph(self, p):
        if p.text.strip() == '':
            return
        p_el = etree.Element('p')
        # TODO: Shouldn't we just append to last element in headstack? What happens after embedded lists?
        if self.current_el is not None:
            if self.current_el.tag == 'div':
                self.current_el.append(p_el)
            else:
                self.current_el.addnext(p_el)
        self.current_el = p_el

    def iterate_runs(self, p):
        '''
        Populates a paragraph level element with its runs properly marked up (these are character level styles)
        Creates a <temp> element to contain the inner XML structure of the paragraph level element
        Iterates over the runs in the paragraph creating "elem" elements sometimes with inner structure.
        These get appended to <temp> which in the end is added to the doc and becomes self.current_el

        [In old converter this was iterateRange (the interateRuns function was not called)]

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

            # if "Heading" in p.style.name:
            #    rtxt = re.sub(r'^[\d\s\.]+', '', rtxt)

            char_style = run.style.name
            is_new_style = True if char_style != last_run_style else False
            # Default Paragraph Font
            if not char_style or char_style == "" or "Default Paragraph Font" in char_style:
                # May be in Default Font but have bold or italic ste
                new_el = getFontElement(run)  # Check for those font characteristics
                if new_el is not None:
                    new_el.text = rtxt
                    temp_el.append(new_el)
                    elem = temp_el.getchildren()[-1]
                elif elem is None:
                    temp_el.text += rtxt
                else:
                    elem.tail += rtxt
            # Footnotes
            elif "footnote" in char_style.lower() or "endnote" in char_style.lower():
                # logging.warning("!!! Must deal with critical edition notes !!!")
                note = self.endnotes.pop(0) if "endnote" in char_style.lower() else self.footnotes.pop(0)
                if elem is None and len(temp_el.getchildren()) > 0:
                    elem = temp_el.getchildren()[-1]
                if elem is not None:
                    back_text = elem.tail if elem.tail and len(elem.tail) > 0 else elem.text
                else:
                    back_text = temp_el.text
                reading = self.process_critical(note, back_text)
                if not reading:
                    rs = note.find('rs')
                    if isinstance(rs, etree._Element):
                        note.remove(rs)
                    temp_el.append(note)
                    elem = note
                elif isinstance(elem, etree._Element):
                    if elem.tag != 'milestone' and (elem.tail is None or len(elem.tail) == 0):
                        elem.text = reading['backtext']
                        elem.append(reading['app'])
                    else:
                        elem.tail = reading['backtext']
                        temp_el.append(reading['app'])
                        elem = reading['app']
                else:
                    temp_el.text = reading['backtext']
                    temp_el.append(reading['app'])
                    elem = reading['app']

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
                    temp_el.text += rtxt
                    if self.debug and char_style not in IGNORABLE_STYLES:
                        print(f"No style definition found for style name, {char_style}: {rtxt}")

                else:
                    if self.debug:
                        logging.debug('Style element {} => {}'.format(char_style, new_el.tag))
                    new_el.text = rtxt
                    temp_el.append(new_el)
                    elem = temp_el.getchildren()[-1]
            else:
                elem.text += rtxt

            last_run_style = char_style
        ### End of iterating runs ###

        if self.current_el is None:
            print('current el is none')  # Should never get here

        # Copy temp_el contents to current_el depending on whether it has children or not
        if self.current_el.tag in ['lb', 'pb', 'milestone']:
            empty_el = self.current_el
            for tmpchld in temp_el.getchildren():
                self.current_el.addnext(tmpchld)
                self.current_el = tmpchld
            empty_el.tail = temp_el.text
            self.current_el = self.current_el.getparent()
            return  # No need to do the following

        curr_child = list(self.current_el) or []
        if len(curr_child) > 0:
            curr_child[-1].tail = temp_el.text
        else:
            self.current_el.text = temp_el.text
        for tempchild in list(temp_el):
            self.current_el.append(tempchild)

        # Deal with numbers at the beginning of headers
        if "heading" in p.style.name.lower():
            headtxt = self.current_el.text
            mtch = re.match(r'^((\d+\.?)+)', headtxt)
            if mtch:
                hnum = mtch.group(1)
                headtxt = headtxt.replace(hnum, '')
                numel = etree.XML('<num>{}</num>'.format(hnum))
                numel.tail = headtxt
                self.current_el.text = ""
                self.current_el.insert(0, numel)

    def process_critical(self, note, bcktxt):
        '''
        Processes footnotes for critical edition with alternate readings from other editions

        :param note:
        :param bcktxt:
        :return:
        '''
        reading = False
        if len(bcktxt) > 0:
            if bcktxt[-1] == '}':
                cestind = bcktxt.rfind('{')
                if cestind > -1:
                    reading = {}
                    lem = bcktxt[cestind + 1:len(bcktxt) - 1].strip()
                    reading['backtext'] = bcktxt[0:cestind]
                    lemed = self.edsig if self.edsig and self.edsig != '' else 'base'
                    lempg = ''
                    notedata = TextConverter.process_critical_note(note)
                    if notedata['lem']:
                        lemed = notedata['lem']['edsigs']
                        lempg = ' n=""'.format(notedata['lem']['edpgs'])

                    app = '<app><lem wit="{}"{}>{}</lem>'.format(lemed, lempg, lem)
                    for vrnt in notedata['variants']:
                        natt = ''
                        if len(vrnt['edpgs']) > 0:
                            natt = ' n="{}"'.format(vrnt['edpgs'])
                        if vrnt['pref']:
                            natt += ' rend="pref"'
                        vartxt = lem if vrnt['txt'] == '' else vrnt['txt']
                        if 'omit' in vartxt:
                            app += '<rdg wit="{}"{} />'.format(vrnt['edsigs'], natt)
                        else:
                            app += '<rdg wit="{}"{}>{}</rdg>'.format(vrnt['edsigs'], natt, vartxt)
                    app += '</app>'
                    reading['app'] = etree.XML(app)
                else:
                    self.mywarning("\n\tFootnote follows close brace as for apparatus, but no preceding open brace " +
                                    "found: {}\n\tNote: {}".format(bcktxt, etree.tostring(note, encoding='unicode')))
        return reading

    @staticmethod
    def process_critical_note(anote):
        notepts = anote.find('rs').text.split(';')
        notedata = {
            'lem': False,
            'variants': []
        }
        for rdg in notepts:
            eds = []
            pref = True if '*' in rdg else False
            rdg = rdg.strip(' *')
            rpts = rdg.split(':')
            tempeds = [ed.strip() for ed in rpts[0].split(',')]
            edsigs = []
            edpgs = []
            for ed in tempeds:
                epts = ed.replace(')', '').split('(')
                edsigs.append(epts[0])
                edpg = epts[1] if len(epts) > 1 else ''
                edpgs.append(edpg)

            edsigs = ' '.join(edsigs)
            edpgs = ' '.join(edpgs)

            rdgtxt = rpts[1] if len(rpts) > 1 else False
            if not rdgtxt:
                notedata['lem'] = {
                    'edsigs': edsigs,
                    'edpgs': edpgs
                }
            else:
                notedata['variants'].append({
                    'pref': pref,
                    'edsigs': edsigs,
                    'edpgs': edpgs,
                    'txt': rdgtxt.strip() if rdgtxt else ''
                })
        return notedata

    def process_multiline_app(self, app_ps):
        orig_current_el = self.current_el
        app_el = etree.XML('<p><app><rdg></rdg></app></p>')
        self.current_el = app_el.find('app').find('rdg')
        for ap in app_ps:
            if ap.text[0] == '{' or ap.text[0] == '}':
                continue
            self.iterate_runs(ap)
            if ap not in app_ps[-2:]:  # Not either the last line of app or the closing }
                lb = etree.XML('<lb/>')
                self.current_el.append(lb)
                self.current_el = lb
        for rn in app_ps[-1].runs:
            char_style = rn.style.name.lower()
            if "footnote" in char_style or "endnote" in char_style:
                note = self.endnotes.pop(0) if "endnote" in char_style.lower() else self.footnotes.pop(0)
                notedata = TextConverter.process_critical_note(note)
                sigla = self.edsig
                pgs = ''
                if notedata['lem']:
                    sigla = notedata['lem']['edsigs']
                    pgs = notedata['lem']['edpgs']
                if self.current_el.tag == 'lb':
                    self.current_el.getparent().set('wit', sigla)
                    if len(pgs) > 0:
                        self.current_el.getparent().set('n', pgs)
                else:
                    self.current_el.set('wit', sigla)
                    if len(pgs) > 0:
                        self.current_el.set('n', pgs)

                for vrnt in notedata['variants']:
                    rdg = '<rdg wit={} n={}>{}</rdg>'.format(vrnt['edsigs'], vrnt['edpgs'], vrnt['txt'])
                    rdg = etree.XML(rdg)
                    app_el.append(rdg)
        orig_current_el.addnext(app_el)
        self.current_el = app_el

    @staticmethod
    def createmilestone(char_style, mstxt):
        mstype = 'line' if 'line' in char_style.lower() else 'page'  # defaults to page
        msnum = mstxt.replace('[', '').replace(']', '')   # Default backup num if regex doesn't match
        mtch = re.match(r'\[?(Page|Line)\s+([^\]]+)\]?', mstxt, re.IGNORECASE)
        if mtch:
            mstype = mtch.group(1)
            msnum = mtch.group(2)
        else:  # In TCD some formatting weirdness in page milestones read as: [21-page Dg]
            mtch = re.match(r'\[?(\d)+\-page\s+([^\]]+)\]?', mstxt, re.IGNORECASE)
            if mtch:
                mstype = 'page'
                msnum = mtch.group(2) + '-' + mtch.group(1)
            # else:s
            #    logging.warning("No match for milestone parts in {}".format(mstxt))
        msel = getStyleElement(char_style)
        msel.set('unit', mstype)
        sep = '.' if '.' in msnum else '-'  # Do we need to check for more separators
        pts     = msnum.split(sep)
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

    def assignids(self):
        print("\rAssigning IDs");
        divs = self.xmlroot.xpath('/*//text//div')
        for divel in divs:
            ancids = list(self.getSiblingPos(divel))
            textpt = ''  # To store the a/b/c for front, body or back
            for anc in divel.iterancestors():  # Iterates ancestors in reverse from current element up to TEI.2
                ancid = self.getSiblingPos(anc)  # returns number for div or a, b, c for front body or back
                if ancid.isnumeric():
                    ancids.append(ancid)
                else:
                    txtpt = ancid
                    break   # break after front body or back to not count text or TEI.2
            ancids.reverse()
            if self.chapnum:
                ancids[0] = self.chapnum
            myid = txtpt + '-'.join(ancids)
            divel.set('id', myid)

    @staticmethod
    def getSiblingPos(el):
        atag = el.tag
        if atag == "front":
            return "a"
        if atag == "body":
            return "b"
        if atag == "back":
            return "c"
        anum = 1
        for sib in el.itersiblings('div', preceding=True):
            anum += 1
        return str(anum)

    def tidyxml(self):
        empty_resp = self.xmlroot.xpath("//publicationStmt/respStmt/name[@n='agent' and not(text())]/parent::*")
        for resp in empty_resp:
            resp.getparent().remove(resp)

    def writexml(self):
        # Determine Name for Resulting XML file
        fname = self.current_file.replace('.docx', '.xml')
        fpth = os.path.join(self.outdir, fname)
        while os.path.isfile(fpth) and not self.overwrite:
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
            genid = self.textid.replace('-text', '')
            mtch = re.search(r'-(\d{4})-', genid)
            fldr = mtch.group(1)[0] if mtch else '0'
            docType = "<!DOCTYPE TEI.2 SYSTEM \"{0}xtib3.dtd\" [ \n" \
                "\t<!ENTITY % thlnotent SYSTEM \"{0}catalog-refs.dtd\" > \n" \
                "\t%thlnotent;\n" \
                "\t<!ENTITY {1} SYSTEM \"../../{2}/{1}-bib.xml\">\n]>".format(self.dtdpath, genid, fldr)

            # Replace profile desc with entity
            pdentity = etree.Entity('thdlprofiledesc')
            pdesc = self.xmlroot.xpath('//profileDesc')[0]
            pdesc.addprevious(pdentity)
            pdesc.getparent().remove(pdesc)

            # Add tibble entity
            tibsrc = etree.XML('<sourceDesc n="tibbibl"></sourceDesc>')
            tibbibl_ent = etree.Entity(genid)
            tibsrc.append(tibbibl_ent)
            docsrc = self.xmlroot.xpath('//sourceDesc')[0]
            docsrc.tail = "\n"
            docsrc.addnext(tibsrc)

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


class ConversionException(Exception):
    pass

