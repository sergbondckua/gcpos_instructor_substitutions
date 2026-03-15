``` cmd windows
python -m PyInstaller --onefile --noconsole --collect-all openpyxl --collect-all pillow --icon=icon.ico main.py --hidden-import=et_xmlfile
```