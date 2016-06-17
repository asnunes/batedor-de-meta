from distutils.core import setup
import py2exe

setup(
    windows=[{'script': 'batedor_de_meta.py', "icon_resources": [(1, "myicon.ico")]}], 
    
    options={
        'py2exe': 
        {
            'includes': ['lxml.etree', 'lxml._elementpath', 'gzip'],
        }
    },
    zipfile=None,
)
