call "C:\home\data\git\jupyter\ppt_macro\venv\Scripts\activate.bat"
cd /d %~dp0
Python src\slideshow_from_drop.py %*
pause