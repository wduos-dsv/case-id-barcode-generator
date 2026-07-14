# Case ID Barcode Generator

## Build
Step 1: Install PyInstaller

```bash
pip install pyinstaller
```

Step 2: Build the .exe File

```bash
PyInstaller --onefile --noconsole qt.py
```

### Adding an Icon
```bash
pyinstaller --onefile --noconsole --icon=my_logo.ico qt.py
```