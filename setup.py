from setuptools import setup, find_packages
import os

def read(*names):
    return open(os.path.join(os.path.dirname(__file__), *names)).read()

setup(
    name="pyvda",
    version=read("version.txt"),
    description="Python implementation of the VirtualDesktopAccessor for manipulating Windows 10 virtual desktops.",
    author="Mike Roberts",
    author_email="mike.roberts.2k10@googlemail.com",
    license="LICENSE.txt",
    url="https://github.com/mrob95/Breathe",
    long_description = read("README.md"),
    long_description_content_type='text/markdown',
    packages=find_packages(exclude=("tests", "tests.*")),
    install_requires=["pywin32"],
    classifiers=[
                   "Environment :: Win32 (MS Windows)",
                   "License :: OSI Approved :: MIT",
                   "Operating System :: Microsoft :: Windows",
                   "Programming Language :: Python :: 2.7",
                   "Programming Language :: Python :: 3.7",
                  ],
)

