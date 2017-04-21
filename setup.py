"""Coati setup module
"""
from setuptools import setup, find_packages

DESCRIPTION = "A tool for programmatically generating PowerPoint reports."
LONG_DESCRIPTION = """
**coati** is a tool that provide a fast simple and highly customizable  way to programmatically 
create PowerPoint presentations. 
coati works using a base PowerPoint template and can replicate and fill the slices with the desire
information. 

Currently coati supports four types of inputs:

 - Excel charts
 - Images (png, svg, jpg)
 - Excel tables
 - Text 
"""
VERSION = "0.8.0"
DISTNAME = "coati"
LICENSE = "MIT"
AUTHOR = "The Intelimetrica Team"
EMAIL = "concato@intelimetrica.com"
URL = "https://github.com/Intelimetrica/coati"
DOWNLOAD_URL = ''

setup(
    name=DISTNAME,
    version=VERSION,
    description=DESCRIPTION,
    long_description=LONG_DESCRIPTION,
    url=URL,
    author=AUTHOR,
    author_email=EMAIL,
    license=LICENSE,

    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Win32 (MS Windows)',
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Build Tools',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2.7',
    ],
    keywords="PowerPoint, automatic reports",
    packages=find_packages(exclude=['pywin32', 'ipython', 'ipython-genutil', 'psycopg2']),
    zip_safe=False,
)
