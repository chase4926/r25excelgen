
from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')

setup(windows=[{'script': 'app.py',
                'icon_resources': [(1, "icon.ico")]}],
            options={"py2exe": {"includes": ["openpyxl", "datetime",
                                             "re", "collections",
                                             "pyglet", "cocos"],
                     #"packages": ["gzip", "lxml", "ssl"],
                     "dll_excludes": ["w9xpopen.exe"],
                     "bundle_files": 3,
                     "compressed": False}},
            zipfile = None)
