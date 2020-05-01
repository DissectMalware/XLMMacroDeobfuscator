from XLMMacroDeobfuscator import __version__

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

entry_points = {
    'console_scripts': [
        'xlmdeobfuscator=XLMMacroDeobfuscator.deobfuscator:main',
        'xlmshell=XLMMacroDeobfuscator.deobfuscator:main_shell'
    ],
}

setup(
    name="XLMMacroDeobfuscator", # Replace with your own username
    version=__version__,
    author="Amirreza Niakanlahiji",
    author_email="aniak2@uis.edu",
    description=(
        "XLMMacroDeobfuscator is a XLM Emulation engine written in Python 3, designed to "
        "analyze and deobfuscate malicious XLM macros, also known as Excel 4.0 macros,"
        "contined in MS Excel files (XLS, XLSM, and XLSB)."),
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/DissectMalware/XLMMacroDeobfuscator",
    packages=["XLMMacroDeobfuscator"],
	entry_points=entry_points,
    license='Apache License 2.0',
    python_requires='>=3.4',
	install_requires=[
        "pyxlsb2",
        "lark-parser", 
        "pywin32",
    ],
)