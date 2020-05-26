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
To deobfuscate macros in Excel documents: 

```
xlmdeobfuscator --file document.xlsm
```

To only get the deobfuscated macros and without any indentation:

```
xlmdeobfuscator --file document.xlsm --no-indent --output-formula-format "[[INT-FORMULA]]"
```

To export the output in JSON format 
```
xlmdeobfuscator --file document.xlsm --export-json result.json
```
To see a sample JSON output, please check [this link](https://pastebin.com/bwmS7mi0) out.
# Command Line (xlmdeobfuscator)

```

          _        _______
|\     /|( \      (       )
( \   / )| (      | () () |
 \ (_) / | |      | || || |
  ) _ (  | |      | |(_)| |
 / ( ) \ | |      | |   | |
( /   \ )| (____/\| )   ( |
|/     \|(_______/|/     \|
   ______   _______  _______  ______   _______           _______  _______  _______ _________ _______  _______
  (  __  \ (  ____ \(  ___  )(  ___ \ (  ____ \|\     /|(  ____ \(  ____ \(  ___  )\__   __/(  ___  )(  ____ )
  | (  \  )| (    \/| (   ) || (   ) )| (    \/| )   ( || (    \/| (    \/| (   ) |   ) (   | (   ) || (    )|
  | |   ) || (__    | |   | || (__/ / | (__    | |   | || (_____ | |      | (___) |   | |   | |   | || (____)|
  | |   | ||  __)   | |   | ||  __ (  |  __)   | |   | |(_____  )| |      |  ___  |   | |   | |   | ||     __)
  | |   ) || (      | |   | || (  \ \ | (      | |   | |      ) || |      | (   ) |   | |   | |   | || (\ (
  | (__/  )| (____/\| (___) || )___) )| )      | (___) |/\____) || (____/\| )   ( |   | |   | (___) || ) \ \__
  (______/ (_______/(_______)|/ \___/ |/       (_______)\_______)(_______/|/     \|   )_(   (_______)|/   \__/

    
XLMMacroDeobfuscator(v 0.1.3) - https://github.com/DissectMalware/XLMMacroDeobfuscator

usage: xlmdeobfuscator [-h] [-f FILE_PATH] [-n] [-x] [-2] [-s] [-d DAY]
                       [--output-formula-format OUTPUT_FORMULA_FORMAT]
                       [--no-indent] [--export-json FILE_PATH]
                       [--start-point CELL_ADDR]

  -h, --help            show this help message and exit
  -f FILE_PATH, --file FILE_PATH
                        The path of a XLSM file
  -n, --noninteractive  Disable interactive shell
  -x, --extract-only    Only extract cells without any emulation
  -s, --start-with-shell
                        Open an XLM shell before interpreting the macros in
                        the input
  -d DAY, --day DAY     Specify the day of month
  --with-ms-excel       Use MS Excel to process XLS files
  --output-formula-format OUTPUT_FORMULA_FORMAT
                        Specify the format for output formulas ([[CELL_ADDR]],
                        [[INT-FORMULA]], and [[STATUS]]
  --no-indent           Do not show indent before formulas
  --export-json FILE_PATH
                        Export the output to JSON
  --start-point CELL_ADDR
                        Start interpretation from a specific cell address
  -2, --no-ms-excel     [Deprecated] Do not use MS Excel to process XLS files

```

# Library
The following example shows how XLMMacroDeobfuscator can be used in a python project to deobfuscate XLM macros:

```python
from XLMMacroDeobfuscator.deobfuscator import process_file

result = process_file(file='path/to/an/excel/file', 
            noninteractive= True, 
            noindent= True, 
            output_formula_format='[[CELL_ADDR]], [[INT-FORMULA]]',
            return_deobfuscated= True)

for record in result:
    print(record)
```

Read requirements.txt to get the list of python libraries that XLMMacroDeobfuscator is dependent on.

You can run XLMMacroDeobfuscator on any OS to extract and deobfuscate macros in xls, xlsm, and xlsb files. No need to install MS Excel.

Note: if you want to use MS Excel (on Windows), you need to install pywin32 library and use --with-ms-excel switch.
If --with-ms-excel is used, xlmdeobfuscator, first, attempts to load xls files with MS Excel, if it fails it uses xlrd2.

\* This code is still heavily under development. Expect to see radical changes in the code.
