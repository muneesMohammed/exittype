## Build on Windows

Install the application dependencies and PyInstaller:

```powershell
py -m pip install -r requirements.txt
py -m pip install pyinstaller
```

Build the windowed, single-file application from the project directory:

```powershell
py -m PyInstaller --noconfirm exittype.spec
```

The executable is created at `dist\exittype.exe`.

The app stores settings and, by default, generated PDFs and JSON files in
`%LOCALAPPDATA%\ExitTypeGenerator`. You can still choose a different output
folder in the app.
