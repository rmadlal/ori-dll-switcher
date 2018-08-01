### Easily switch to a different dll in Ori DE. ###

This program will locate the requested dll and copy it into `...\Ori DE\oriDE_Data\Managed` with the name `Assembly-CSharp.dll` for you.  
dll names will be saved for future use.

Download `dll_switcher.py` and run it via the console with a dll name as an argument. You can place this file anywhere (I'd put it in the Ori directory)

Examples:  
1. `python dll_switcher.py rando` --> an "Open file" window will open. Choose your rando dll. 
2. `python dll_switcher.py nopause` --> an "Open file" window will open. Choose your nopause dll. 
3. `python dll_switcher.py rando` --> your rando dll was remembered from the first time and will be automatically copied in.
