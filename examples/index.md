---
layout: page
title: "Examples"
---

## Examples

Check out how easily xlwings-powered spreadsheets can be distributed by trying out these examples:

### Downloads

**Example 1: Fibonacci Sequence**

* **Lite:** [fibonacci.zip][] (32.5 KB)
* **Standalone:** [fibonacci_standalone.zip][] (6.3 MB)

[fibonacci.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/fibonacci.zip
[fibonacci_standalone.zip]: https://bitbucket.org/zoomeranalytics/xlwings_examples/downloads/fibonacci_standalone.zip

Dependencies of the Lite Version: Python, pywin32

**Example 2: Coming up...**

### Lite Versions

These versions are small in size but require an installation of Python. It is highly recommended to install one of the
following scientific Python distributions as they already contain all of the necessary packages used in the examples,
most importantly pywin32, numpy, scipy and pandas.

* [Anaconda](https://store.continuum.io/cshop/anaconda/)
* [WinPython](http://winpython.sourceforge.net/) (see Notes below)
* [Canopy](https://www.enthought.com/products/canopy/)
* [Python(x,y)](https://code.google.com/p/pythonxy/)


### Standalone Versions

These versions run out-of-the-box after unzipping without any dependencies but are bigger in size.


### Instructions

1. Download the zip-file
2. Right-Click > Extract All... > Extract
3. Open the Spreadsheet in the unzipped folder
4. "Protected View": click on "Enable Editing"
5. Optional: If Excel gives you an additional "Security Warning": click on "Enable Content", then
   please close and reopen the file
6. Run the examples by clicking on the "Run" button

### Notes

**WinPython**: Since WinPython doesn't change the PATH environment variables, you either have to add the location
  of the Python interpreter to the PATH manually or change the directory in the spreadsheet as follows:

* Press `Alt-F11` to fire up the VBA Editor
* Double-click the `xlwings` module
* In the `RunPython` function, change `PYTHON_DIR = ""` to the directory of where `python.exe` is, e.g.:
`PYTHON_DIR = "C:\WinPython-64bit-2.7.6.3\python-2.7.6"`
