
:: generates the version.txt file
C:\Users\chick\AppData\Local\Programs\Python\Python311\Scripts\create-version-file metadata.yml --outfile version.txt
C:\Users\chick\AppData\Local\Programs\Python\Python311\Scripts\pyinstaller --noconsole --onefile --version-file version.txt create_opsheet.py