# Proposal Manager

*Created using the [openpyxl](https://pypi.org/project/openpyxl/), [py2app](https://pypi.org/project/py2app/), and [tkinter](https://docs.python.org/3/library/tkinter.html) python libraries.*

The app was curated for Redbud Development, INC. 

## Project Description

The Proposal Manager desktop app accepts a xlsx Excel file using the 'Create Proposal' button and processes data from the first sheet within the file, which is based on a budget proposal Excel template. 

When the program finishes, Excel will open and a finalized and ready-to-print budget proposal will be inserted to the left of the original budget sheet. 

## Setting up the desktop app

The app runs from a tkninter-based desktop app that can be accessed 

>Setting up the app requires the [py2app](https://pypi.org/project/py2app/) library, which utilizes Python 3+ and can be run easily within a pipenv shell.

*Installing pipenv and activating a shell:*
```
brew update
brew install pipenv 
```
*Within main directory (proposal-manager) run:*
```
pipenv shell
```
*Install dependencies:*
```
pipenv update 
**OR**
python install -r requirements.txt
```
*Navigate to the 'proposal-manager' directory and run the following command to create the app:*
```
python setup.py py2app
```
*The app can be created with an icns image fileby instead running:*
```
py2applet --make-setup proposal_manger_gui.py *.icns
python setup.py py2app
```
*Clean build directory:*
```
rm -rf build dist
```

- More detail on creating the app with the pyapp library can be found at [https://pypi.org/project/py2app/](https://pypi.org/project/py2app/)
- The app will appear in a dist/ directory within the proposal-manager folder, but can be moved to anywhere on the machine.

## Repository Organization

This repository contains the scripts used to parse and create the budget proposal, a python script for the tkinter integration of the app, and examples images of the app and desktop icon. 

- build/
  - Contains the app once it has been created 
- dist/
  - Contains build files, can be deleted. 
- Pipfile
  - Program dependency requirements as managed by the pipenv shell
- requirements.txt
  - Program requirements created by ``` pip freeze > requirements.txt ```
- proposal_creator.py 
  - Script that contains backend code for processing excel data using the [openpyxl](https://pypi.org/project/openpyxl/) library.
- proposal_manager_gui.py
  - Script that uses [tkinter](https://docs.python.org/3/library/tkinter.html) to create a graphical interface for the app. 
- setup.py
  - Script created via py2applet and used to build the app.  

