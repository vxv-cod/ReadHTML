pyinstaller -w -F -i "logo.ico" ReadHTML.py
xcopy %CD%\*.xltx %CD%\dist /H /Y /C /R
xcopy %CD%\*.xlsx %CD%\dist /H /Y /C /R