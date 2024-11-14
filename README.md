<img src="https://github.com/user-attachments/assets/372ae9d3-072f-4047-9494-400fa06bcbcc" alt="App Icon" width="200" align="center"><br><br>This application gives 4j School District employees a convenient way to fill in mileage expense reports. It was created by Rebecca Medley on the Mentor Team.<br><br><img src="https://github.com/user-attachments/assets/ed0cae14-58e3-4a0a-8708-d47d20a1d0ba" alt="Example GUI" width="220" align="right" style="margin-right: 15px; margin-bottom: 15px;"><br>

---

### Features
 - Uses 'official' distances between buildings
 - A vanilla copy of the expense report PDF is bundled with the application and a new PDF is created every time it's filled in -- this precludes annoying errors that can happen when the same PDF is edited multiple times.
 - Auto-complete is enabled for school and purpose columns
 - Documents can be opened and saved to allow for incremental progress before creating the final PDF.

### Releases
The [current macOS version](https://github.com/inductivekickback/smiles/releases/) is tested on Big Sur and newer (Intel and Apple silicon). It's not a problem if your IT department prevents you from dragging things to the Applications folder -- drag this to your Desktop and run it from there instead.

### Building
This application relies on the [pyqt6](https://pypi.org/project/PyQt6/), [pymupdf](https://pypi.org/project/PyMuPDF/), and [pyinstaller](https://pypi.org/project/pyinstaller/) projects. It was built with Python 3.8 but later Python versions are probably fine. The recommended process is to clone the repository and then:
```
$ python3 -m venv smiles/env
$ cd smiles
$ source env/bin/activate
$ pip install --upgrade pip
$ pip install -r requirements.txt
```
It can be run directly from the command line. Specify an optional document to open when starting:
```
$ python3 smiles.py PATH-TO-DOCUMENT
```
Use pyinstaller to build the macOS bundle:
```
$ python3 pyinstaller smiles.spec
```
