# -*- coding: utf-8 -*-
from distutils.core import setup
import py2exe
includes = ["email","xlrd"]
options = {"py2exe":
               {
                   "compressed":1,
                   "optimize":2,
                   "includes":includes,
                   "bundle_files":1,
               }
        }
setup(
    version ="3.0.1",
    description= 'baobiao zhizuo',
    options = options,
    zipfile = None,
    console = [{"script":"excel_test.py","icon_resources":[(1,"biao.ico")]}]
)

