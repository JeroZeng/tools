from distutils.core import setup
import py2exe

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