# -*- coding: utf-8 -*-
"""
Created on Mon May 07 09:22:56 2018

@author: Ozgur Yarikkas
"""

from distutils.core import setup
import py2exe, sys, os
from glob import glob
from distutils.filelist import findall
import matplotlib

sys.setrecursionlimit(5000)
sys.argv.append('py2exe')

matplotlibdatadir = matplotlib.get_data_path()
matplotlibdata = findall(matplotlibdatadir)
matplotlibdata_files = []
for f in matplotlibdata:
    dirname = os.path.join('matplotlibdata', f[len(matplotlibdatadir)+1:])
    matplotlibdata_files.append((os.path.split(dirname)[0], [f]))

mpld = matplotlib.get_py2exe_datafiles()

# change this!
DATA = [('imageformats',['C:\\Python27/Lib/site-packages/PyQt4/plugins/imageformats/qgif4.dll'])]

SETUP_DICT = {
    'windows': [{
        'script': 'GuiRun.py',
    }],

    'zipfile': None,
#    r'mpl-data'
    'data_files': matplotlibdata_files,

    'data_files': (
        ('', glob(r'C:\Windows\SYSTEM32\msvcp100.dll')),
        ('', glob(r'C:\Windows\SYSTEM32\msvcr100.dll')),
    ),
    'data_files': mpld,
    'data_files': DATA,

    'options': {
        'py2exe': {
            'bundle_files': 3,
            'compressed': 2,
            'optimize': 2,
            'includes': ['sip', 'PyQt4.QtCore'
             , 'PyQt4'
             , 'PyQt4.QtGui'
             , 'PyQt4.Qt'
             ,'reportlab.rl_settings',
             'win32com'
             ,'win32com.client'
             , 'matplotlib'
             , 'matplotlib.backends'
             , 'matplotlib.backends.backend_qt4agg'
             , 'matplotlib.figure'],
             'excludes': ['nbformat','win32com.gen_py',"six.moves.urllib.parse",
            '_gtkagg', '_tkagg', 'wxagg', '_agg2',
            '_cairo', '_cocoaagg',
            '_fltkagg', '_gtk', '_gtkcairo',
#            'email',
#            'bsddb', 'curses',
#            'pywin.debugger','pywin.debugger.dbgcon', 'pywin.dialogs',
#                                'Tkconstants', 'Tkinter',
#                                'doctest', 'test', 'sqlite3'
                                ],
            'xref': False,
            'skip_archive': False,
            'ascii': False,
            'custom_boot_script': '',
            'dist_dir': 'dist',  # Put .exe in dist/
#            'packages': ['pytz'],
            'dll_excludes': ['libgdk-win32-2.0-0.dll', 'libgobject-2.0-0.dll'
#                                'tcl84.dll', 'tk84.dll'
#                             'msvcr71.dll', 'w9xpopen.exe',
#                                     'API-MS-Win-Core-LocalRegistry-L1-1-0.dll',
#                                     'API-MS-Win-Core-ProcessThreads-L1-1-0.dll',
#                                     'API-MS-Win-Security-Base-L1-1-0.dll',
#                                     'KERNELBASE.dll',
#                                     'POWRPROF.dll'
#                                     ' MSVCP90.dll'
                                     ]
#                                     ['libgdk-win32-2.0-0.dll', 'libgobject-2.0-0.dll',
#                             'msvcr71.dll', 'w9xpopen.exe',
#                                     'API-MS-Win-Core-LocalRegistry-L1-1-0.dll',
#                                     'API-MS-Win-Core-ProcessThreads-L1-1-0.dll',
#                                     'API-MS-Win-Security-Base-L1-1-0.dll',
#                                     'KERNELBASE.dll',
#                                     'POWRPROF.dll']
        },
    }
}

setup(**SETUP_DICT)
