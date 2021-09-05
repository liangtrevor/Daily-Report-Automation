# Daily-Report-Automation

The purpose of this script is to automatically move data from a GenLive POS report into a formatted .xlsx report. 

# Getting Started

To get started, first open the command prompt from Windows search (type cmd) and take note of its default **path**. 

![image](https://user-images.githubusercontent.com/88129677/132111480-73acbe8c-37fc-4efc-bc89-82e633bd66e5.png)

Then download [Python 3](https://www.python.org/downloads/) on your computer. 

To make things easier, make sure Python is installed on the default path you see in cmd so you do not have to change the directory. Python should be installed in a directory with a path similar to this:

![image](https://user-images.githubusercontent.com/88129677/132111308-ea6f73e0-81d4-4ab5-8887-39e6aecd689b.png)

In cmd, enter ```pip install openpyxl==2.6.2```. 

If openpyxl installs with no issues, congrats! You have all the dependencies to make this work. 

# To Use

To use this script, put a POS report saved in .xlsx format, the .xlsx daily report you wish to output to, and main.py in the **same folder**. 

Once you have that set up, run the script by right clicking and selecting "Open with Python". It should open with Python by default but you can make sure by specifying. 

![image](https://user-images.githubusercontent.com/88129677/132111433-5f182a62-287f-4f29-b317-637d44ffa614.png)

You will be prompted with on screen instructions. Once the program finishes executing, a copy of the daily report will appear in the folder the script was placed in. 
