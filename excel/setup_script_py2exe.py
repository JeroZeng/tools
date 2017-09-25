from distutils.core import setup
import py2exe

# to use tkinter
'''
py2exe.build_exe.py2exe.old_prepare = py2exe.build_exe.py2exe.plat_prepare
def new_prep(self):
  self.old_prepare()
  from _tkinter import TK_VERSION, TCL_VERSION
  self.dlls_in_exedir.append('tcl{0}.dll'.format(TCL_VERSION.replace('.','')))
  self.dlls_in_exedir.append('tk{0}.dll'.format(TK_VERSION.replace('.','')))
py2exe.build_exe.py2exe.plat_prepare = new_prep
'''

setup(
    options = {'py2exe': {'bundle_files': 1}},
    console = [{
            "script": "merge.py",
            "icon_resources": [(1, "excel.ico")]
            }],
    zipfile = None
)

# usage: python setup_script_py2exe.py py2exe
# list options for py2exe
# http://www.py2exe.org/index.cgi/ListOfOptions