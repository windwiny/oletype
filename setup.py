# -*- coding: utf-8 -*-
from setuptools import setup, find_packages

try:
    long_description = open("README.md").read()
except IOError:
    long_description = ""

setup(
    name="oletype",
    version="0.3.0",
    description="generate win32com excel object classes pyi file",
    license="MIT",
    url="https://github.com/windwiny/oletype",
    author="windwiny",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[],
    long_description=long_description,
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Python :: 3.8",
    ],
    keywords=[
        'win32com',
        'excel.application'
    ],
    data_files=[
        ('a', ['oletype/excel.pyi']),
    ],

)
