cd /d %~dp0
@rem call "venv\Scripts\activate.bat"
@rem args[1]: 0 or 1; 0 -> slide_aspect_ratio = 4 / 3, 1 -> slide_aspect_ratio = 16 / 9
@rem args[2]: row number of grid
@rem args[3]: column number of grid
@rem args[4]: 0 or 1; 0 -> mode = 'fill', 1 -> mode = 'fit'
Python src\slideshow_from_drop.py 1 1 1 1 %*
@rem pause