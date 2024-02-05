# Workorder to Opsheet
Generates a set of op sheets from a bom xlsx generated by solidworks, sorted by the first operation for operators to be able to sort operations by material.

Opens a dialog box where the bom xlsx file can be selected, and then saves a pdfto user's downloads, named after the bom file.

Ex.
```
bom file: 9929-06-100_A.xlsx
output pdf: 9929-06-100_A-opsheet.pdf
```

## Packaging
Packaged using [pyinstaller](https://pyinstaller.org/en/stable/index.html), ends up in `dist/create_opsheet`.

Windows: simply navigate to the root of this project and run
```
run.bat
```

### Generating metadata
It also uses [pyinstaller-versionfile](https://pypi.org/project/pyinstaller-versionfile/)
 to generate a version file that can then be passed into the pyinstaller cli command with the `--version-file` param.

[metadata.yml](metadata.yml) contains the default info, and the block in the python file overwrites any defaults.



## Changelog

### 1.1.0
Feb 2, 2024
 - Add material table for saw
 - Convert inches to feet
 - Add 'WaterJet', now 4 op codes
 - Add metadata generation using pyinstaller-versionfile

### 1.0.0
January 14, 2024
First try.