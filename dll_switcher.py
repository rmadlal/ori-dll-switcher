import os
import sys
import argparse
import json
import shutil
from tkinter import filedialog

ORI_ROOT = r'C:\Program Files (x86)\Steam\steamapps\common\Ori DE'
ASSEMBLY_CSHARP = os.path.join(ORI_ROOT, 'oriDE_Data', 'Managed', 'Assembly-CSharp.dll')
DLL_NAMES = 'dll_names.json'

try:
    with open(DLL_NAMES, 'r') as f:
        dll_names = json.load(f)
except OSError:
    dll_names = {}


def store_dll_names():
    with open(DLL_NAMES, 'w') as f:
        json.dump(dll_names, f)


def validate_ori_root():
    ori_root = ORI_ROOT
    assembly_csharp = ASSEMBLY_CSHARP
    if not os.path.exists(ori_root):
        path = filedialog.askdirectory(mustexist=True,
                                       title=f'Select the Ori DE directory')
        if not path:
            sys.exit('Ori DE directory not found')
        ori_root = os.path.abspath(path)
        assembly_csharp = os.path.join(ori_root, 'oriDE_Data', 'Managed', 'Assembly-CSharp.dll')
    return ori_root, assembly_csharp


def get_dll_path(dll_name, ori_root, force_open_dialog):
    if not force_open_dialog and dll_name in dll_names:
        dll_path = dll_names[dll_name]
    else:
        path = filedialog.askopenfilename(filetypes=[('DLL', '*.dll')],
                                          initialdir=ori_root,
                                          title=f'Choose "{dll_name}" dll file')
        if not path:
            sys.exit('dll file not found')
        dll_path = os.path.abspath(path)
        dll_names[dll_name] = dll_path
        store_dll_names()
    return dll_path


def main():
    parser = argparse.ArgumentParser(description='Switch to a different Ori dll.')
    parser.add_argument('dll', help="the dll's name")
    parser.add_argument('-o', '--open',
                        action='store_true', help='force "Open file" dialog window', dest='force_open')
    args = parser.parse_args()

    ori_root, assembly_csharp = validate_ori_root()

    dll_name = args.dll
    dll_path = get_dll_path(dll_name, ori_root, args.force_open)
    if not os.path.exists(dll_path):
        dll_path = get_dll_path(dll_name, ori_root, True)

    shutil.copy(dll_path, assembly_csharp)

    print('Done!')


if __name__ == '__main__':
    main()
