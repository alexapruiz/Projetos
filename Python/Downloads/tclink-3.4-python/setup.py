"""
Python Distutils Script for TCLink

distutils documentation
http://www.python.org/doc/current/dist

TClink
http://www.trustcommerce.com/tclink.html

$Id: setup.py,v 1.3 2003/07/18 23:39:28 adam Exp $
"""

import os
from distutils.core import setup, Extension


# if you have ssl/crypto libraries installed in non standard
# locations you may need to specify their location here
include_dirs = ['/usr/local/ssl', '/usr/local/openssl'] 


tclink_version="3.4"
machine_info = "%s %s"%( os.uname()[0], # sysname
                         os.uname()[-1] # machine
                         )

machine_info = machine_info.replace(' ', '-')
pytclink_version = '"%s-Python-%s"'%(tclink_version,machine_info)


tclink_extension = Extension("tclink",
                             ["py_tclink.c"],
                             libraries=["ssl", "crypto"],
                             define_macros=[('TCLINK_VERSION',pytclink_version)],
                             include_dirs = include_dirs
                             )


setup(
    name="TCLink",
    version=tclink_version,
    description="TrustCommerce Transaction Client Library",
    author="TrustCommerce",
    author_email="developer@trustcommerce.com",
    url="http://www.trustcommerce.com/tclink.php",
    ext_modules=[ tclink_extension ]
    )
    
