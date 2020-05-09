# XLMMacroDeobfuscator
XLMMacroDeobfuscator can be used to decode obfuscated XLM macros (also known as Excel 4.0 macros). It utilizes an internal XLM emulator to interpret the macros, without fully performing the code.

It supports both xls, xlsm, and xlsb formats. 

It uses [xlrd2](https://github.com/DissectMalware/xlrd2), [pyxlsb2](https://github.com/DissectMalware/pyxlsb2) and its own parser to extract cells and other information from xls, xlsb and xlsm files, respectively.

You can also find XLM grammar in xlm-macro-en.lark

# Installing the emulator

1. Install using pip

```
pip install XLMMacroDeobfuscator
```

2. Installing the latest development

```
pip install -U https://github.com/DissectMalware/XLMMacroDeobfuscator/archive/master.zip
```

# Running the emulator
To run the script 

```
xlmdeobfuscator --file document.xlsm
```

# Usage

```
usage: xlmdeobfuscator [-h] [-f FILE] [-n] [-x] [-2] [-s]

optional arguments:
  -h, --help              show this help message and exit
  -f FILE, --file FILE    The path of a XLSM file
  -n, --noninteractive    Disable interactive shell
  -x, --extract-only      Only extract cells without any emulation
  -2, --no-ms-excel       Do not use MS Excel to process XLS files
  -s, --start-with-shell  Open an XLM shell before interpreting the macros in
                          the input
```

Read requirements.txt to get the list of python libraries that XLMMacroDeobfuscator is dependent on.

You can run XLMMacroDeobfuscator on any OS to extract and deobfuscate macros in xls, xlsm, and xlsb files. No need to install MS Excel.

Note: if you want to use MS Excel (on Windows), you need to install pywin32 library. if you do not want to use MS Excel, use --no-ms-excel.
Otherwise, xlmdeobfuscator, first, attempts to load xls files with MS Excel, if it fails it uses xlrd2.

\* This code is still heavily under development. Expect to see radical changes in the code.
