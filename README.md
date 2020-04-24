# XLMMacroDeobfuscator
XLMMacroDeobfuscator can be used to decode obfuscated XLM macros (also known as Excel 4.0 macros). It utilizes an internal XLM emulator to interpret the macros, without fully performing the code.

It supports both xls, xlsm, and xlsb formats. 

It uses pyxlsb2 and its own parser to extract cells and other information from xlsb and xlsm files. However, it relies on MS Excel to extract such information. As such, you need to have MS Excel on the machine if you want to process xls files.

Note: Processing xlsm and xlsb files are much faster than xls files (in two orders of magnitude)

Soon, an xls parser will be included to make it independent of MS Excel

WARNING: tmp\tmp.zip contains real malicious excel documents (password: infected). Please only run them in a testing environment.

# Running the script
To run the script 

```
python  XLMMacroDeobfuscator.py --file document.xlsm
```

# Prerequisit
To parse xlsb file, XLMMacroObfuscator relies on [pyxlsb2](https://github.com/DissectMalware/pyxlsb2). To install the pyxlsb2 library:

```
pip install -U https://github.com/DissectMalware/pyxlsb2/releases/download/0.0.2/pyxlsb2-0.0.2-py3-none-any.whl
```

\* This code is heavily under development. Expect to see radical changes in the code.
