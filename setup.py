from XLMMacroDeobfuscator import __version__
import os

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

project_dir = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(project_dir, 'README.md')) as f:
    long_description = f.read()

entry_points = {
    'console_scripts': [
        'xlmdeobfuscator=XLMMacroDeobfuscator.deobfuscator:main',
    ],
}

setup(
    name="XLMMacroDeobfuscator",
    version=__version__,
    author="Amirreza Niakanlahiji",
    author_email="aniak2@uis.edu",
    description=(
        "XLMMacroDeobfuscator is an XLM Emulation engine written in Python 3, designed to "
        "analyze and deobfuscate malicious XLM macros, also known as Excel 4.0 macros, "
        "contained in MS Excel files (XLS, XLSM, and XLSB)."),
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
        "xlrd2",
        "untangle==1.1.1"
    ],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: Information Technology",
        "Intended Audience :: Science/Research",
        "Intended Audience :: System Administrators",
        "License :: OSI Approved :: Apache Software License",
        "Natural Language :: English",
        "Operating System :: OS Independent",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.4",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Topic :: Security",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    package_data={'XLMMacroDeobfuscator':['xlm-macro.lark.template', 'configs/get_workspace.conf']},
)
