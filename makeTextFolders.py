#!env/bin/python

from os import listdir, makedirs
from os.path import exists, join
from shutil import move
import re

DIRPATH = './workspace/out'
PATT = r'^[^-]+-(\d{4})'

files = [fnm for fnm in listdir(DIRPATH) if fnm.endswith('.xml')]

for fnm in files:
    # print(fnm)
    mtch = re.search(PATT, fnm)
    if mtch:
        fldrnm = join(DIRPATH, mtch.group(1))
        if not exists(fldrnm):
            makedirs(fldrnm)
            src = join(DIRPATH, fnm)
            dest = join(fldrnm, fnm)
            move(src, dest)


print(f"Created text folders and moved texts into them in: {DIRPATH}")
