from setuptools import setup, find_packages
import os
import pyvda

def read(*names):
    return open(os.path.join(os.path.dirname(__file__), *names)).read()

setup(
    name="pyvda",
    version=pyvda.__version__,
    description="Python implementation of the VirtualDesktopAccessor for manipulating Windows 10 virtual desktops.",
    author="Mike Roberts",
    author_email="mike.roberts.2k10@googlemail.com",
    license="LICENSE.txt",
    url="https://github.com/mrob95/py-VirtualDesktopAccessor",
    long_description = read("README.md"),
    long_description_content_type='text/markdown',
    packages=find_packages(exclude=("tests", "tests.*")),
    install_requires=["pywin32", "comtypes"],
    classifiers=[
                   "Environment :: Win32 (MS Windows)",
                   "License :: OSI Approved :: MIT License",
                   "Operating System :: Microsoft :: Windows",
                   "Programming Language :: Python :: 3.7",
                  ],
)

