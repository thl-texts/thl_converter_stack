#!env/bin/python

from os import listdir, makedirs
from os.path import exists, join, split
from shutil import move
import re

DIRPATH = './workspace/out'
PATT = r'^[^-]+-(\d{4})'
DTD = '<?xml version="1.0" encoding="UTF-8"?>' \
        '<!DOCTYPE TEI.2 SYSTEM "../../../../../xml/dtds/xtib3.dtd" [' \
        '<!ENTITY % thlnotent SYSTEM "../../../catalog-refs.dtd" >' \
        '%thlnotent;' \
        ' <!ENTITY lccw-TNUM SYSTEM "../../0/lccw-TNUM-bib.xml">' \
      ']><TEI.2'
SRCPATT = '<sourceDesc (.*)</sourceDesc>'
SRCNEW = ' <sourceDesc n="tibbibl">&lccw-TNUM;</sourceDesc>'


def cleanup(src, dest, tnum):
    with open(src, 'r', encoding='utf-8') as filein:
        data = filein.read()
        pts = data.split('<TEI.2')
        newdata = DTD.replace('TNUM', tnum) + pts[1]

        ndpts = newdata.split('<sourceDesc')
        allnew = ndpts[0] + SRCNEW.replace('TNUM', tnum)
        ndpts2 = ndpts[1].split('</sourceDesc>')
        allnew = allnew + ndpts2[1]

    [fpth, fnm] = split(dest)
    if not exists(fpth):
        makedirs(fpth)
    with open(dest, 'w', encoding='utf-8') as fileout:
        fileout.write(allnew)


def main():
    flsindir = listdir(DIRPATH)
    files = [fnm for fnm in listdir(DIRPATH) if fnm.endswith('.xml')]
    for fnm in files:
        # print(fnm)
        mtch = re.search(PATT, fnm)
        if mtch:
            tnum = mtch.group(1)
            fldrnm = join(DIRPATH, tnum)
            src = join(DIRPATH, fnm)
            dest = join(fldrnm, fnm)
            cleanup(src, dest, tnum)

    print(f"Created text folders and moved texts into them in: {DIRPATH}")


if __name__ == '__main__':
    main()
