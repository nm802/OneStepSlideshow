cd /d %~dp0
@rem call "venv\Scripts\activate.bat"
Python src\slideshow_from_drop.py %*
@rem pause