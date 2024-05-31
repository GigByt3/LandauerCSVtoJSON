

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

## My errors
* libgssapi_krb5
message
```commandline
Traceback (most recent call last):
  File "/home/michael/Code/LandauerCSVtoJSON/main.py", line 14, in <module>
    from PyQt6.QtQml import QQmlApplicationEngine
ImportError: libgssapi_krb5.so.2: cannot open shared object file: No such file or directory
```
solutions
* remove import for that modules
```commandline
sudo apt-get install libgssapi-krb5-2
```

In dialog
* No JVM shared library file (libjvm.so) found. Try setting up the JAVA_HOME environment variable properly
  https://www.baeldung.com/find-java-home

Maybe try [camalot](https://github.com/camelot-dev/camelot/blob/master/README.md)

Other reading
* https://github.com/actions/setup-python
* https://docs.github.com/en/actions/automating-builds-and-tests/building-and-testing-python
* https://www.pythonguis.com/tutorials/packaging-pyqt5-applications-pyinstaller-macos-dmg/
* https://packaging.python.org/en/latest/tutorials/packaging-projects/