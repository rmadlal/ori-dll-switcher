## Easily switch to a different DLL in Ori DE. ##

### [Download](https://github.com/rmadlal/ori-dll-switcher/releases) ###

![Screenshot](/screenshot.png)

This program will locate the requested DLL and copy it into `...\Ori DE\oriDE_Data\Managed` with the name `Assembly-CSharp.dll` for you.

Run OriDLLSwitcher.exe and choose a short name for your DLL (e.g "rando", "nopause" etc). Then click on "Set DLL" and select the proper file. Finally, hit "Apply".  
DLL names will be saved so that you would simply pick your DLL from the list and click "Apply".

You can also create a shortcut for an instant DLL switch by selecting a DLL name and clicking on "Create shortcut".  
Note: running the program with a command line argument of the DLL name will do the same as running the shortcut. (instant DLL switch, no GUI)

The program loads much faster when running the script directly (dll_switcher.py instead of OriDLLSwitcher.exe). Python 3 must be installed.  
Run `pip install -r requirements.txt` followed by `python dll_switcher.py`.
