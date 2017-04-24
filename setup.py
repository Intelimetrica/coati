#!/usr/bin/env python
# -*- encoding: utf-8 -*-
from __future__ import absolute_import
from __future__ import print_function

import io
import re
from glob import glob
from os.path import basename
from os.path import dirname
from os.path import join
from os.path import splitext

from setuptools import find_packages
from setuptools import setup


def read(*names, **kwargs):
    return io.open(
        join(dirname(__file__), *names),
        encoding=kwargs.get('encoding', 'utf8')
    ).read()


setup(
    name='coati',
    version='0.1.0',
    license='MIT',
    description='A tool for programmatically generating PowerPoint reports.',
    long_description="""
        **coati** is a tool that provides a fast, simple and highly customizable  way to programmatically
        generate PowerPoint presentations.
        coati works using a base PowerPoint template and can replicate and fill the slides with the desire
        information.

        Currently coati supports four types of inputs:

         - Excel charts
         - Images (png, svg, jpg)
         - Excel tables
         - Text
        """,
    author='Intelimetrica team',
    author_email='mailrelay@intelimetrica.com',
    url='https://github.com/Intelimetrica/coati',
    packages=find_packages('src'),
    package_data={'coati': ['src/coati/templates/*.txt']},
    package_dir={'': 'src'},
    py_modules=[splitext(basename(path))[0] for path in glob('src/*.py')],
    include_package_data=True,
    zip_safe=False,
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: Implementation :: PyPy',
        'Topic :: Software Development :: Build Tools',
    ],
    keywords=[
        "PowerPoint, automatic reports",
    ],
    install_requires=[
        'python-dotenv==0.4.0',
        'colorlog==2.10.0'
    ],
    extras_require={
        # eg:
        #   'rst': ['docutils>=0.11'],
        #   ':python_version=="2.6"': ['argparse'],
    },
    entry_points={
        'console_scripts': [
            'coati = coati.cli:main',
        ]
    },
)
