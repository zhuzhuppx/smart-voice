pyinstaller -F -w -i dancer.ico main.py
cxfreeze main.py --target-dir build --base-name=Win32GUI