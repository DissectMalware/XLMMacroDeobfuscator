# XLMMacroDeobfuscator
XLMMacroDeobfuscator can be used to decode obfuscated XLM macros (also known as Excel 4.0 macros). It utilizes an internal XLM emulator to interpret the macros, without fully performing the code.

It supports both xls, xlsm, and xlsb formats. 

It uses [xlrd2](https://github.com/DissectMalware/xlrd2), [pyxlsb2](https://github.com/DissectMalware/pyxlsb2) and its own parser to extract cells and other information from xls, xlsb and xlsm files, respectively.

You can also find XLM grammar in [xlm-macro-lark.template](XLMMacroDeobfuscator/xlm-macro.lark.template)

# Installing the emulator

1. Install using pip

```
pip install XLMMacroDeobfuscator
```

2. Installing the latest development

```
pip install -U https://github.com/DissectMalware/xlrd2/archive/master.zip
pip install -U https://github.com/DissectMalware/pyxlsb2/archive/master.zip
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

To use a config file
```
xlmdeobfuscator --file document.xlsm -c default.config
```

default.config file must be a valid json file, such as:

```json
{
	"no-indent": true,
	"output-formula-format": "[[CELL-ADDR]] [[INT-FORMULA]]",
	"non-interactive": true,
	"output-level": 1
}
```

# Command Line 

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

    
XLMMacroDeobfuscator(v0.1.7) - https://github.com/DissectMalware/XLMMacroDeobfuscator

usage: deobfuscator.py [-h] [-c FILE_PATH] [-f FILE_PATH] [-n] [-x] [-2]
                       [--with-ms-excel] [-s] [-d DAY]
                       [--output-formula-format OUTPUT_FORMULA_FORMAT]
                       [--no-indent] [--export-json FILE_PATH]
                       [--start-point CELL_ADDR] [-p PASSWORD]
                       [-o OUTPUT_LEVEL]

optional arguments:
  -h, --help            show this help message and exit
  -c FILE_PATH, --config_file FILE_PATH
                        Specify a config file (must be a valid JSON file)
  -f FILE_PATH, --file FILE_PATH
                        The path of a XLSM file
  -n, --noninteractive  Disable interactive shell
  -x, --extract-only    Only extract cells without any emulation
  -2, --no-ms-excel     [Deprecated] Do not use MS Excel to process XLS files
  --with-ms-excel       Use MS Excel to process XLS files
  -s, --start-with-shell
                        Open an XLM shell before interpreting the macros in
                        the input
  -d DAY, --day DAY     Specify the day of month
  --output-formula-format OUTPUT_FORMULA_FORMAT
                        Specify the format for output formulas ([[CELL-ADDR]],
                        [[INT-FORMULA]], and [[STATUS]]
  --no-indent           Do not show indent before formulas
  --export-json FILE_PATH
                        Export the output to JSON
  --start-point CELL_ADDR
                        Start interpretation from a specific cell address
  -p PASSWORD, --password PASSWORD
                        Password to decrypt the protected document
  -o OUTPUT_LEVEL, --output-level OUTPUT_LEVEL
                        Set the level of details to be shown (0:all commands,
                        1: commands no jump 2:important commands 3:strings in
                        important commands).
  --timeout N           stop emulation after N seconds (0: not interruption
                        N>0: stop emulation after N seconds)
```

# Library
The following example shows how XLMMacroDeobfuscator can be used in a python project to deobfuscate XLM macros:

```python
from XLMMacroDeobfuscator.deobfuscator import process_file

result = process_file(file='path/to/an/excel/file', 
            noninteractive= True, 
            noindent= True, 
            output_formula_format='[[CELL_ADDR]], [[INT-FORMULA]]',
            return_deobfuscated= True,
            timeout= 30)

for record in result:
    print(record)
```

* note: the xlmdeofuscator logo will not be shown when you use it as a library

# Requirements

Please read requirements.txt to get the list of python libraries that XLMMacroDeobfuscator is dependent on.

xlmdeobfuscator can be executed on any OS to extract and deobfuscate macros in xls, xlsm, and xlsb files. You do not need to install MS Excel.

Note: if you want to use MS Excel (on Windows), you need to install pywin32 library and use --with-ms-excel switch.
If --with-ms-excel is used, xlmdeobfuscator, first, attempts to load xls files with MS Excel, if it fails it uses [xlrd2 library](https://github.com/DissectMalware/xlrd2).

# Project Using XLMMacroDeofuscator
XLMMacroDeofuscator is adopted in the following projects:
* [CAPE Sandbox](https://github.com/ctxis/CAPE)
* [FAME](https://certsocietegenerale.github.io/fame/)
* [REMNUX](https://remnux.org/)
* [IntlOwl](https://github.com/certego/IntelOwl)
* [Assemblyline 4](https://cybercentrecanada.github.io/assemblyline4_docs/) by Canadian Centre for Cyber Security 

Please contact me if you incorporated XLMMacroDeofuscator in your project.

# How to Contribute
If you found a bug or would like to suggest an improvement, please create a new issue on the [issues page](https://github.com/DissectMalware/XLMMacroDeobfuscator/issues).

Feel free to contribute to the project forking the project and submitting a pull request.

You can reach [me (@DissectMlaware) on Twitter](https://twitter.com/DissectMalware) via a direct message.

