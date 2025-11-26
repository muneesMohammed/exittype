pyinstaller --onefile --icon=exittype.ico --name=exittype label.py  -w



pyinstaller --onefile --windowed --icon=exittype.ico --name=exittype label.py --add-data "icons;icons"

pyinstaller --onefile --windowed --icon=exittype.ico --name=exittype --add-data "icons;icons" label.py

