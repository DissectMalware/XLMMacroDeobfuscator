# XLMMacroDeobfuscator
XLMMacroDeobfuscator can be used to decode obfuscated XLM macros (also known as Excel 4.0 macros). It utilizes an internal XLM emulator to interpret the macros, without fully performing the code.

It supports both xls, xlsm, and xlsb formats. 

It uses [pyxlsb2](https://github.com/DissectMalware/pyxlsb2) and its own parser to extract cells and other information from xlsb and xlsm files. However, it relies on MS Excel to extract such information. As such, you need to have MS Excel on the machine if you want to process xls files.

Note: Processing xlsm and xlsb files are much faster than xls files (in two orders of magnitude)

Soon, an xls parser will be included to make it independent of MS Excel

WARNING: tmp\tmp.zip contains real malicious excel documents (password: infected). Please only run them in a testing environment.

You can also find XLM grammar in xlm-macro.lark

# Running the script
To run the script 

```
python  deobfuscator.py --file document.xlsm
```

# Usage

```
usage: deobfuscator.py [-h] [-f FILE] [-n] [-x] [-s]

optional arguments:
  -h, --help              show this help message and exit
  -f FILE, --file FILE    The path of a XLSM file
  -n, --noninteractive    Disable interactive shell
  -x, --extract-only      Only extract cells without any emulation
  -s, --start-with-shell  Open an XLM shell before interpreting the macros in
                          the input
```

# Prerequisit
To parse xlsb file, XLMMacroObfuscator relies on [pyxlsb2](https://github.com/DissectMalware/pyxlsb2). To install the pyxlsb2 library:

```
pip install -U pyxlsb2
```

It also requires Microsoft Excel in order to process XLS files. However, if only XLSM or XLSB files are being processed, MS Excel is not needed.

\* This code is heavily under development. Expect to see radical changes in the code.
