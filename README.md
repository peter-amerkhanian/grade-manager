# About
This project spawned out of a coworker's need to move a large 
amount of data from an Excel spreadsheet into an odd spreadsheet
app that did not allow multi-cell copy and paste from excel. I used a 
combination of xlwings and pyautogui to move the data over.

While this has limited use outside of my specific work situation,
parts of the code may be useful to someone else facing the challenge
of moving large amounts of data into an app that does not allow for
multi-cell pasting.

## Install / Run
Python 3.6 and Windows recommended. Navigate to the directory of the 
repo and run the following:
```
$ pip3 install -r requirements.txt
$ python3 grade_manager.py
```

## Mac specific
One may need to reinstall some 
of the requirements and also install xCode.  

it wil all look more like this:
```
$ pip3 install pyobjc-core
$ pip3 install pyobjc
$ pip3 install xlwings # here you'll be prompted to download xCode
$ pip3 install pyautogui
$ pip3 install pyperclip
$ pip3 install -r requirements.txt
```

Instead of running in the terminal, you can
just double-click:  
grade_manager.command

