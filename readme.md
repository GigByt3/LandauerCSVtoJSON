Build As .app
```shell
pyinstaller --name 'Landauer Report Processor' \
    --windowed \
    --icon='GUI/icons/apple.icns' \
    --add-data='GUI/icons/16x16.png:.' \
    --add-data='GUI/icons/24x24.png:.' \
    --add-data='GUI/icons/32x32.png:.' \
    --add-data='GUI/icons/48x48.png:.' \
    --add-data='GUI/icons/256x256.png:.' \
    --add-data='requirements.txt:.' \
    main.py
```
Build As .dmg
```shell
hdiutil create -volname "LandauerReportProcessor" \
    -srcfolder "/Users/.../dist/Landauer Report Processor.app" \
    -ov \
    -format UDZO "LandauerReportProcessor.dmg"
```