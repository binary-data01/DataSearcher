del /Q /F /S dist
pyinstaller -F -w -n data_searcher main.py
"C:\Program Files\WinRAR\WinRAR.exe" a dist\data_searcher.zip dist\data_searcher.exe
pause