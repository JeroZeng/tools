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

# list options for py2exe
# http://www.py2exe.org/index.cgi/ListOfOptions