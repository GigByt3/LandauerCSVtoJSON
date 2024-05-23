

Build As .app
```shell
pyinstaller --name 'Landauer Report Processor' \
    --icon='GUI/icons/apple.icns' --windowed \
    --add-data='GUI/icons/16x16.png:.' \
    --add-data='GUI/icons/24x24.png:.' \
    --add-data='GUI/icons/32x32.png:.' \
    --add-data='GUI/icons/48x48.png:.' \
    --add-data='GUI/icons/256x256.png:.' \
    --add-data='requirements.txt:.' \
    --onefile \
    main.py
```
Build As .dmg
```
```