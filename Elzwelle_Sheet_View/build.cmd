copy ..\elzwelle_sheet_view.py .
mkdir dist
mkdir dist\elzwelle_sheet_view
\opt\miniconda3\Scripts\pyinstaller.exe elzwelle_sheet_view.py
copy \opt\miniconda3\Library\bin\libcrypto-3-x64.dll 	dist\elzwelle_sheet_view\_internal
copy \opt\miniconda3\Library\bin\libssh2.dll 			dist\elzwelle_sheet_view\_internal
copy \opt\miniconda3\Library\bin\libssl-3-x64.dll		dist\elzwelle_sheet_view\_internal
