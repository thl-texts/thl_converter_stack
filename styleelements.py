

########## Python Library to Convert Style Names to Elements ##########

from lxml import etree
import json

##### Global Element Dictionary ######

# "keydict" is a dictionary of keys string matched with arrays of word style names
# The keys are the same keys in the element dictionary
# The values are an array of word styles names that should use that key to look up their element stats
# The function "createStyleKeyDict" creates a flat dictionary with keys of Word Style names matched with the "keys"
# used to look up the xml element specs (tagname and attributes)
keydict = {
    "abbr": ["Abbreviation"],
    "add-by-ed": ["Added by Editor"],
    "annotations": ["Annotations"],
    "auth-chi": ["author Chinese"],
    "auth-eng": ["author English"],
    "auth-gen": ["X-Author Generic"],
    "auth-ind": ["X-Author Indian"],
    "auth-san": ["author Sanskrit"],
    "auth-tib": ["X-Author Tibetan", "author Tibetan"],
    "author": ["author"],
    "date-range": ["Date", "Date Range", "Dates", "X-Dates"],
    "dates": ["X-Dates"],
    "dox-cat": ["Doxographical-Bibliographical Category", "X-Doxo-Biblio Category"],
    "emph-strong": ["Emphasis Strong", "X-Emphasis Strong", "Strong"],
    "emph-weak": ["Emphasis Weak", "Subtle Emphasis1", "X-Emphasis Weak"],
    # "endn-char": ["Endnote Characters"],
    # "endn-ref": ["endnote reference"],
    # "endn-text": ["endnote text"],
    # "endn-text-char": ["Endnote Text Char"],
    "epithet": ["Epithet"],
    # "foot-bibl": ["Footnote Bibliography"],
    # "foot-char": ["Footnote Characters"],
    # "foot-ref": ["footnote reference"],
    # "foot-text": ["footnote text"],
    "hyperlink": ["Hyperlink", "FollowedHyperlink"],
    "illegible": ["Illegible"],
    "lang-chi": ["Lang Chinese"],
    "lang-eng": ["Lang English"],
    "lang-fre": ["Lang French"],
    "lang-ger": ["Lang German"],
    "lang-jap": ["Lang Japanese"],
    "lang-kor": ["Lang Korean"],
    "lang-mon": ["Lang Mongolian"],
    "lang-nep": ["Lang Nepali"],
    "lang-pali": ["Lang Pali"],
    "lang-sans": ["Lang Sanskrit"],
    "lang-span": ["Lang Spanish"],
    "lang-tib": ["Lang Tibetan"],
    "line-num": ["line number", "LineNumber"],
    "line-num-dig": ["Line Number,digital"],
    "line-num-print": ["Line Number Print"],
    "line-num-tib": ["Line Number Tib"],
    "mantra": ["X-Mantra"],
    "monuments": ["Monuments", "X-Monuments"],
    "name-bud-deity": ["X-Name Buddhist  Deity", "Name Buddhist  Deity"],
    "name-bud-deity-coll": ["X-Name Buddhist Deity Collective", "Name Buddhist Deity Collective"],
    "name-ethnic": ["X-Name Ethnicity", "Name of ethnicity"],
    "name-fest": ["Name festival", "X-Name Festival"],
    "name-gen": ["X-Name Generic", "Name generic"],
    "name-org": ["X-Name Organization", "Name organization"],
    "name-org-clan": ["X-Name Clan", "Name org clan"],
    "name-org-line": ["X-Name Lineage", "Name org lineage"],
    "name-org-monastery": ["X-Name Monastery", "Name organization monastery"],
    "name-pers-human": ["X-Name Personal Human", "Name Personal Human"],
    "name-pers-other": ["X-Name Personal Other", "Name Personal Other"],
    "name-place": ["X-Name Place", "Name Place"],
    "name-rel-pract": ["X-Religious Practice", "Religious practice", "Name ritual"],
    "page-num": ["PageNumber", "page number"],
    "page-num-print-ed": ["Page Number Print Edition"],
    "pages": ["Pages"],
    # "placeholder-text": ["Placeholder Text1"],
    "plain-text": ["Plain Text"],
    "pub-place": ["publication place"],
    "publisher": ["publisher"],
    "root-text": ["Root Text", "Root text"],
    "sa-bcad": ["Sa bcad"],
    "speak-bud-deity": ["X-Speaker Buddhist Deity", "SpeakerBuddhistDeity", "Epithet Buddhist Deity", "Speaker Epithet Buddhist Deity"],
    "speak-bud-deity-coll": ["Speaker Buddhist Deity Collective"],
    "speak-human": ["X-Speaker Human", "SpeakerHuman"],
    "speak-other": ["SpeakerOther", "X-Speaker Other"],
    "speak-unknown": ["X-Speaker Unknown"],
    "speak-gene": ["Speaker generic"],
    "speech-inline": ["Speech Inline"],
    "term-chi": ["X-Term Chinese", "term Chinese"],
    "term-eng": ["term English", "X-Term English"],
    "term-fre": ["term French"],
    "term-ger": ["term German"],
    "term-jap": ["term Japanese"],
    "term-kor": ["term Korean"],
    "term-mon": ["term Mongolian", "X-Term Mongolian"],
    "term-nep": ["term Nepali"],
    "term-pali": ["term Pali", "X-Term Pali"],
    "term-sans": ["X-Term Sanskrit", "term Sanskrit"],
    "term-span": ["term Spanish"],
    "term-tech": ["X-Term Technical"],
    "term-tib": ["X-Term Tibetan", "term Tibetan"],
    "text-group": ["TextGroup", "X-Text Group"],
    "text-title": ["Text Title"],
    "text-title-san": ["Text Title Sanksrit"],
    "text-title-tib": ["Text Title Tibetan"],
    "tib-line-number": ["TibLineNumber"],
    "title": ["Title"],
    "title-chap": ["Title of Chapter", "Colophon Chapter Title", "Text Title in Chapter Colophon"],
    "title-cite-other": ["Title in Citing Other Texts"],
    "title-own-non-tib": ["Title (Own) Non-Tibetan Language"],
    "title-own-tib": ["Title (Own) Tibetan", "Colophon Text Titlle", "Colophon Text Title", "Text Title in Colophon"],
    "title-section": ["Title of Section"],
    "unclear": ["Unclear"]
}

# Defining global dictionary for Word style to element keys to be populated by createStyleKeyDict() function
styledict = {}

# Elements is a keyed dictionary of information for defining XML elements
elements = {
    "abbr": {
        "tag": "abbr",
        "attributes": {"expan": ""},
    },
    "add-by-ed": {
        "tag": "add",
        "attributes": {"n": "editor"},
    },
    "annotations": {
        "tag": "note",
        "attributes": {"type": "annotation"},
    },
    "auth-chi": {
        "tag": "persName",
        "attributes": {"type": "author", "n": "chinese"},
    },
    "auth-eng": {
        "tag": "persName",
        "attributes": {"type": "author", "n": "engish"},
    },
    "auth-gen": {
        "tag": "persName",
        "attributes": {"type": "author", "n": "generic"},
    },
    "auth-ind": {
        "tag": "persName",
        "attributes": {"type": "author", "n": "indian"},
    },
    "auth-san": {
        "tag": "persName",
        "attributes": {"type": "author", "n": "sanskrit"},
    },
    "auth-tib": {
        "tag": "persName",
        "attributes": {"type": "author", "n": "tibetan"},
    },
    "author": {
        "tag": "persName",
        "attributes": {"type": "author"},
    },
    "date-range": {  # need to get converter to recognize split as what to split on and markup accordingly.
        "tag": "dateRange",
        "attributes": {"from": "", "to": ""},
        "split": "-",
        "childels": "date",
    },
    "dox-cat": {
        "tag": "term",
        "attributes": {"type": "doxcat"},
    },
    "emph-strong": {
        "tag": "hi",
        "attributes": {"rend": "strong"},
    },
    "emph-weak": {
        "tag": "hi",
        "attributes": {"rend": "weak"},
    },
    "epithet": {
        "tag": "name",
        "attributes": {"type": "epithet"},
    },
    "hyperlink": {
        "tag": "xref",
        "attributes": {"n": "%0%", "type": "url"},
    },
    "illegible": {
        "tag": "gap",
        "attributes": {"n": "", "reason": "illegible"},
    },
    "lang-chi": {
        "tag": "seg",
        "attributes": {"lang": "chi"},
    },
    "lang-eng": {
        "tag": "seg",
        "attributes": {"lang": "eng"},
    },
    "lang-fre": {
        "tag": "seg",
        "attributes": {"lang": "fre"},
    },
    "lang-ger": {
        "tag": "seg",
        "attributes": {"lang": "ger"},
    },
    "lang-jap": {
        "tag": "seg",
        "attributes": {"lang": "jap"},
    },
    "lang-kor": {
        "tag": "seg",
        "attributes": {"lang": "kor"},
    },
    "lang-mon": {
        "tag": "seg",
        "attributes": {"lang": "mon"},
    },
    "lang-nep": {
        "tag": "seg",
        "attributes": {"lang": "nep"},
    },
    "lang-pali": {
        "tag": "seg",
        "attributes": {"lang": "pli"},
    },
    "lang-sans": {
        "tag": "seg",
        "attributes": {"lang": "san"},
    },
    "lang-span": {
        "tag": "seg",
        "attributes": {"lang": "spa"},
    },
    "lang-tib": {
        "tag": "seg",
        "attributes": {"lang": "tib"},
    },
    "line-num": {
        "tag": "milestone",
        "attributes": {"unit": "line", "n": "%TXT%"},
    },
    "line-num-dig": {
        "tag": "milestone",
        "attributes": {"unit": "digline"},
    },
    "line-num-print": {
        "tag": "milestone",
        "attributes": {"unit": "line", "n":"%TXT%"},
    },
    "line-num-tib": {
        "tag": "milestone",
        "attributes": {"unit": "tibline", "n":"%TXT%"},
    },
    "mantra": {
        "tag": "rs",
        "attributes": {"type": "mantra"},
    },
    "monuments": {
        "tag": "placeName",
        "attributes": {"type": "monument"},
    },
    "name-bud-deity": {
        "tag": "persName",
        "attributes": {"type": "bud-deity"},
    },
    "name-bud-deity-coll": {
        "tag": "orgName",
        "attributes": {"type": "bud-deity"},
    },
    "name-ethnic": {
        "tag": "orgName",
        "attributes": {"type": "ethnicity"},
    },
    "name-fest": {
        "tag": "term",
        "attributes": {"type": "festival"},
    },
    "name-gen": {
        "tag": "persName",
        "attributes": {"type": "generic"},
    },
    "name-org": {
        "tag": "orgName",
        "attributes": {},
    },
    "name-org-clan": {
        "tag": "orgName",
        "attributes": {"type": "clan"},
    },
    "name-org-line": {
        "tag": "orgName",
        "attributes": {"type": "lineage"},
    },
    "name-org-monastery": {
        "tag": "placeName",
        "attributes": {"type": "monastery"},
    },
    "name-pers-human": {
        "tag": "persName",
        "attributes": {"type": "human"},
    },
    "name-pers-other": {
        "tag": "persName",
        "attributes": {"type": "other"},
    },
    "name-place": {
        "tag": "placeName",
        "attributes": {},
    },
    "name-rel-pract": {
        "tag": "term",
        "attributes": {"type": "religious-practice"},
    },
    "page-num": {
        "tag": "milestone",
        "attributes": {"unit": "page", "n": "%TXT%"},
    },
    "page-num-print-ed": {
        "tag": "milestone",
        "attributes": {"unit": "page", "n": "%TXT%"},
    },
    "pages": {
        "tag": "num",
        "attributes": {"type": "page-range"},
    },
    "plain-text": {
        "tag": "hi",
        "attributes": {"rend": "plain"},
    },
    "pub-place": {
        "tag": "pubPlace",
        "attributes": {},
    },
    "publisher": {
        "tag": "publisher",
        "attributes": {},
    },
    "root-text": {
        "tag": "seg",
        "attributes": {"type": "roottext"},
    },
    "sa-bcad": {
        "tag": "rs",
        "attributes": {"type": "sabcad"},
    },
    "speak-bud-deity": {
        "tag": "persName",
        "attributes": {"type": "speaker-bud-deity"},
    },
    "speak-bud-deity-coll": {
        "tag": "persName",
        "attributes": {"type": "speaker-bud-deity-coll"},
    },
    "speak-human": {
        "tag": "persName",
        "attributes": {"type": "speaker-human"},
    },
    "speak-other": {
        "tag": "persName",
        "attributes": {"type": "speaker-other"},
    },
    "speak-unknown": {
        "tag": "persName",
        "attributes": {"type": "speaker-unknown"},
    },
    "speak-gene": {
        "tag": "persName",
        "attributes": {"type": "speaker-generic"},
    },
    "speech-inline": {
        "tag": "q",
        "attributes": {"rend": "inline"},
    },
    "term-chi": {
        "tag": "term",
        "attributes": {"lang": "chi"},
    },
    "term-eng": {
        "tag": "term",
        "attributes": {"lang": "eng"},
    },
    "term-fre": {
        "tag": "term",
        "attributes": {"lang": "fre"},
    },
    "term-ger": {
        "tag": "term",
        "attributes": {"lang": "ger"},
    },
    "term-jap": {
        "tag": "term",
        "attributes": {"lang": "jap"},
    },
    "term-kor": {
        "tag": "term",
        "attributes": {"lang": "kor"},
    },
    "term-mon": {
        "tag": "term",
        "attributes": {"lang": "mon"},
    },
    "term-nep": {
        "tag": "term",
        "attributes": {"lang": "nep"},
    },
    "term-pali": {
        "tag": "term",
        "attributes": {"lang": "pli"},
    },
    "term-sans": {
        "tag": "term",
        "attributes": {"lang": "san"},
    },
    "term-span": {
        "tag": "term",
        "attributes": {"lang": "spa"},
    },
    "term-tech": {
        "tag": "term",
        "attributes": {"n": "technical"},
    },
    "term-tib": {
        "tag": "term",
        "attributes": {"lang": "tib"},
    },
    "text-group": {
        "tag": "title",
        "attributes": {"level": "s", "n": "text-group"},
    },
    "text-title": {
        "tag": "title",
        "attributes": {"level": "m"},
    },
    "text-title-san": {
        "tag": "title",
        "attributes": {"level": "m", "lang": "san"},
    },
    "text-title-tib": {
        "tag": "title",
        "attributes": {"level": "m", "lang": "tib"},
    },
    "tib-line-number": {
        "tag": "milestone",
        "attributes": {"unit": "line", "lang": "tib"},
    },
    "title": {
        "tag": "title",
        "attributes": {},
    },
    "title-chap": {
        "tag": "title",
        "attributes": {"level": "a", "n": "chapter", "type": "internal"},
    },
    "title-cite-other": {
        "tag": "title",
        "attributes": {"level": "m", "type": "external"},
    },
    "title-own-non-tib": {
        "tag": "title",
        "attributes": {"level": "m", "n": "non-tib", "type": "internal"},
    },
    "title-own-tib": {
        "tag": "title",
        "attributes": {"level": "m", "lang": "tib", "type": "internal"},
    },
    "title-section": {
        "tag": "title",
        "attributes": {"level": "a", "n": "section", "type": "internal"},
    },
    "unclear": {
        "tag": "unclear",
        "attributes": {},
    }
}


def getStyleElement(style_name):
    '''
    Returns the XML element object for a particular style name
    :param style_name:
    :return:
    '''
    elemdef = getStyleTagDef(style_name)
    if elemdef is None:
        print("Character style name {} was not found.".format(style_name))
        return etree.XML('<s n="{}"></s>'.format(style_name))
    elem = etree.Element(elemdef['tag'])
    if 'attributes' in elemdef:
        atts = elemdef['attributes']
        if isinstance(atts, dict):
            for nm, val in atts.items():
                elem.set(nm, val)
    return elem


def getStyleTagDef(style_name):
    '''
    Returns the definition of the tag as a python dictionary with "tag" and "attributes" keys
    The style_name parameter can be the name of a style or the key in the element dictionary.
    :param style_name:
    :return:
    '''
    global elements, styledict
    # Create styledict if not already created
    if len(styledict) == 0:
        styledict = createStyleKeyDict()

    # check if style_name is key or if we need to look the key up in the styledict
    if style_name in elements:
        stkey = style_name
    elif style_name in styledict:
        stkey = styledict[style_name]
    else:
        stkey = False

    # return the element def it there
    if stkey in elements:
        return elements[stkey]
    #otherwise return none
    return None


def getTagFromStyle(style_name):
    '''
    This function returns the string version of the xml tag with information filled out
    The style_name parameter can be the name of a style or the key in the element dictionary.
    If it is the style name, the element key is looked up in the style-key dictionary
    :param style_name:
    :return:
    '''
    global elements, styledict
    eldef = getStyleTagDef(style_name)
    elout = "<{0}".format(eldef['tag'])
    for att in eldef['attributes']:
        elout += ' {0}="{1}"'.format(att, eldef['attributes'][att])
        elout += '></{0}>'.format(eldef['tag'])
    return elout


def createStyleKeyDict(tolower=False):
    """
    Creates a dictionary keyed on Word Style name that returns the key for the univeral element array and stores in a global
    Returns the global if it's already populated. The dictionary returned is keyed on Word style name and returns the
    universal element key to use in the Element dictionary. This way more than one Word Style can have the same markup.
    The initial key dict has as its key the key to the element dictionary and as its values arrays of Word Style names.

    :param tolower: whether or not to lowercase the Word Style names used for keys in this dictionary
    :return: styledict: The flat one-to-one dictionary of Word Style Names (capitalized or all lower) and Element dictionary keys.
                        This can then be used to look up the Element definition for any Word Style Names
    """
    global keydict, styledict

    if len(styledict) == 0:
        for k in keydict.keys():
            styles = keydict[k]
            for stnm in styles:
                skey = stnm
                if tolower:
                    skey = skey.lower()
                styledict[skey] = k

    return styledict


def buildElement(stynm, text=None, vals=list()):
    el = getTagFromStyle(stynm)
    if text:
        if '%TXT%' in el:
            el = el.replace('%TXT%', text)
        elif '></' in el:
            el = el.replace('></', '>{0}</'.format(text))

    for n, v in vals:
        el = el.replace('%{0}%'.format(n), v)

    return el


def list_all():
    '''
    Lists all styles / keys/ elements in dictionary
    '''
    skl = createStyleKeyDict()
    keys = skl.keys()
    keys.sort()
    for k in keys:
        el = getTagFromStyle(k)
        print("{0:<40} :\t\t{1:<25}:\t\t{2}".format(k, skl[k], el))
        # print "%s\t\t:\t\t%s" % (k, skl[k])


def main():
    skl = createStyleKeyDict()
    myst = 'Doxographical-Bibliographical Category'
    if myst in skl:
        mu = buildElement(myst)
        print("{0} : {1}".format(myst, mu))
    else:
        print("{0} ain't in no dictionary".format(myst))

    myel = getStyleTagDef(myst)

    print("Tag name is: {0}".format(myel['tag']))

    # Lists all styles / keys/ elements in dictionary
    # list_all()

    #print json.dumps(skl, indent=4)


if __name__ == "__main__":
    main()

