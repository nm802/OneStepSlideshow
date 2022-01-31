cd /d %~dp0
call "venv\Scripts\activate.bat"
Python src\slideshow_from_drop.py %*
pause