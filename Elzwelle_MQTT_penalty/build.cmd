copy ..\elzwelle_mqtt_penalty.py .
mkdir dist
mkdir dist\elzwelle_mqtt_penalty
\opt\miniconda3\Scripts\pyinstaller.exe elzwelle_mqtt_penalty.py
copy \opt\miniconda3\Library\bin\libcrypto-3-x64.dll dist\elzwelle_mqtt_penalty\_internal
copy \opt\miniconda3\Library\bin\libssh2.dll dist\elzwelle_mqtt_penalty\_internal
