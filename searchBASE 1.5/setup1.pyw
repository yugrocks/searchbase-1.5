#A script to compile the source code of searchBASE using py2exe

from distutils.core import setup
import py2exe



setup(
    windows=[
        {
            "script":"searchBASE.pyw",
            
        }
    ],
    
)
