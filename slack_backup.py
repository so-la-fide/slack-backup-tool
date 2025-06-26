# .github/workflows/build-exe.yml
name: Build Slack Backup EXE

on:
  push:
    branches: [ main ]
  workflow_dispatch:

jobs:
  build-windows:
    runs-on: windows-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pyinstaller requests
    
    - name: Build EXE
      run: |
        pyinstaller --onefile --windowed --name="SlackBackup" --icon=icon.ico slack_backup.py
    
    - name: Upload EXE
      uses: actions/upload-artifact@v3
      with:
        name: SlackBackup-Windows
        path: dist/SlackBackup.exe

  build-mac:
    runs-on: macos-latest
    
    steps:
    - uses: actions/checkout@v3
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pyinstaller requests
    
    - name: Build Mac App
      run: |
        pyinstaller --onefile --windowed --name="SlackBackup" --icon=icon.icns slack_backup.py
    
    - name: Upload Mac App
      uses: actions/upload-artifact@v3
      with:
        name: SlackBackup-macOS
        path: dist/SlackBackup
