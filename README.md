# XLMMacroDeobfuscator
XLMMacroDeobfuscator can be used to decode obfuscated XLM macros (also known as Excel 4.0 macros). It utilizes an internal XLM emulator to interpret the macros, without fully performing the code.

It supports both xls and xlsm formats. 

It uses its own parser to extract cells and other information from xlsm files. However, it relies on MS Excel to extract such information. As such, you need to have MS Excel on the machine if you want to process XLS files.

Note: Processing XLSM files are much faster than XLS files (in two orders of magnitude)

Soon, an XLS parser will be included to make it independent of MS Excel

# Running the script
To run the script 

```
python  XLMMacroDeobfuscator.py --file document.xlsm
```

* This code is heavily under development. Expect to see radical changes in the code.
