import argparse
import json
import os
import shutil
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox

from win32com.client import Dispatch

TITLE = 'Ori DLL Switcher'
ORI_ROOT = r'C:\Program Files (x86)\Steam\steamapps\common\Ori DE'
ASSEMBLY_CSHARP = r'oriDE_Data\Managed\Assembly-CSharp.dll'
SWITCHER_JSON = 'dll_switcher.json'
KEY_ROOT = 'root'
KEY_DLL_NAMES = 'dll_names'


class SwitcherException(Exception):

    def __init__(self, message):
        super(SwitcherException, self).__init__(message)


class Switcher(object):

    def __init__(self):
        self._data = self._load_data_json() or self._init_data()

    @staticmethod
    def _init_data():
        return {
            KEY_ROOT: ORI_ROOT,
            KEY_DLL_NAMES: {}
        }

    @staticmethod
    def _check_data_valid(data):
        return KEY_ROOT in data and KEY_DLL_NAMES in data

    def _load_data_json(self):
        try:
            with open(SWITCHER_JSON) as f:
                data = json.load(f)
                if not self._check_data_valid(data):
                    return None
                return data
        except OSError:
            return None

    def save_data_json(self):
        with open(SWITCHER_JSON, 'w') as f:
            json.dump(self._data, f)

    @staticmethod
    def get_assemblycsharp_path(ori_root):
        path = os.path.join(ori_root, ASSEMBLY_CSHARP)
        return path if os.path.exists(os.path.dirname(path)) else None

    @property
    def ori_root(self):
        return self._data[KEY_ROOT]

    @ori_root.setter
    def ori_root(self, ori_root):
        self._data[KEY_ROOT] = ori_root

    @property
    def assembly_csharp(self):
        return self.get_assemblycsharp_path(self.ori_root)

    @property
    def dll_names(self):
        return self._data[KEY_DLL_NAMES]

    def add_dll_name(self, name):
        self.set_dll_path(name, '')

    def get_dll_path(self, name):
        return self.dll_names[name] if name in self.dll_names else None

    def set_dll_path(self, name, path):
        # Check if this is the original-location Assembly-CSharp.dll
        if path == self.assembly_csharp:
            raise SwitcherException('This file cannot be chosen')
        self.dll_names[name] = path

    def switch(self, name):
        path = self.get_dll_path(name)
        if not path:
            raise SwitcherException(f'DLL not found for "{name}"')
        shutil.copy(path, self.assembly_csharp)


class Ui(object):
    def __init__(self, switcher: Switcher):
        self._switcher = switcher
        self._tk_root = Tk()
        self._tk_root.title(TITLE)
        self._tk_root.columnconfigure(0, weight=1)
        self._tk_root.rowconfigure(1, weight=1)
        self._tk_root.geometry('260x90')
        self._tk_root.protocol('WM_DELETE_WINDOW', self._on_destroy)

        # UI elements

        self._combobox = Combobox(self._tk_root, values=tuple(self._switcher.dll_names), state="readonly")
        self._combobox.grid(row=0, column=0, columnspan=2, padx=8, pady=8, sticky=W + E)
        self._combobox.bind('<<ComboboxSelected>>', lambda evt: self._update_chosen_dll())
        self._combobox.bind('<Return>', lambda evt: self._add_new_dll_name(self._combobox.get()))
        if self._switcher.dll_names:
            self._combobox.current(0)

        self._button_new = Button(self._tk_root, text='New', command=self._button_new)
        self._button_new.grid(row=0, column=2, padx=8, sticky=W + E)

        self._button_locate = Button(self._tk_root, text='Locate DLL', command=self._locate_dll)
        self._button_locate.grid(row=1, column=0, padx=8, pady=8, sticky=W + S)

        self._button_create_shortcut = Button(
            self._tk_root, text='Create shortcut', command=self._create_shortcut, state=DISABLED)
        self._button_create_shortcut.grid(row=1, column=1, padx=8, pady=8, sticky=S)

        self._button_apply = Button(self._tk_root, text='Apply', command=self._apply, state=DISABLED)
        self._button_apply.grid(row=1, column=2, padx=8, pady=8, sticky=E + S)

        self._validate_ori_root()
        self._update_chosen_dll()

    def _validate_ori_root(self):
        if self._switcher.assembly_csharp:
            return

        messagebox.showinfo(message='Please locate your Ori DE directory', parent=self._tk_root, title=TITLE)
        while True:
            ori_root = filedialog.askdirectory(parent=self._tk_root,
                                               title='Select the Ori DE directory',
                                               mustexist=True)
            if not ori_root:
                messagebox.showerror(message='Ori DE directory not found', parent=self._tk_root, title=TITLE)
                sys.exit(1)

            ori_root = os.path.abspath(ori_root)
            if self._switcher.get_assemblycsharp_path(ori_root):
                break

            messagebox.showerror(
                message='Invalid Ori DE directory. Should be the directory that contains oriDE.exe',
                parent=self._tk_root, title=TITLE)

        self._switcher.ori_root = ori_root

    # new dll path was set
    def _validate_path(self, dll_path):
        self._button_apply.configure(state=NORMAL if dll_path else DISABLED)
        self._button_create_shortcut.configure(state=NORMAL if dll_path else DISABLED)

    # user chose a name from the list
    def _update_chosen_dll(self):
        dll_name = self._combobox.get()
        dll_path = self._switcher.get_dll_path(dll_name)
        self._validate_path(dll_path)
        self._button_locate.configure(state=NORMAL if dll_name else DISABLED)

    # user clicked on button_new
    def _button_new(self):
        self._combobox.configure(state=NORMAL)
        self._combobox.set('')
        self._combobox.focus()
        self._update_chosen_dll()

    # user pressed enter after typing a new name
    def _add_new_dll_name(self, name):
        if not name:
            return
        self._switcher.add_dll_name(name)
        self._combobox.configure(values=tuple(self._switcher.dll_names), state="readonly")
        self._update_chosen_dll()

    # user clicked on button_locate
    def _locate_dll(self):
        dll_name = self._combobox.get()
        path = filedialog.askopenfilename(filetypes=[('DLL', '*.dll')],
                                          initialdir=self._switcher.ori_root,
                                          parent=self._tk_root,
                                          title=f'Choose your "{dll_name}" DLL file')
        if not path:
            return

        try:
            dll_path = os.path.abspath(path)
            self._switcher.set_dll_path(dll_name, dll_path)
            self._validate_path(dll_path)
        except SwitcherException as e:
            messagebox.showerror(message=e, title=TITLE)

    # user clicked on button_apply
    def _apply(self):
        try:
            dll_name = self._combobox.get()
            self._switcher.switch(dll_name)
            messagebox.showinfo(message=f'Switched to {dll_name} DLL', parent=self._tk_root, title=TITLE)
            self._on_destroy()
        except SwitcherException as e:
            # Should never happen
            messagebox.showerror(message=e, title=TITLE)

    # user clicked on button_create_shortcut
    def _create_shortcut(self):
        dll_name = self._combobox.get()
        shortcut_name = f'SwitchTo{dll_name.capitalize()}.lnk'
        win_cmd_path = r'C:\Windows\System32\cmd.exe'
        program_arg = os.path.basename(sys.argv[0])

        if program_arg.endswith('.py'):
            program_arg = 'python ' + program_arg
        elif program_arg.endswith('.exe'):
            program_arg = program_arg[:-4]
        else:
            messagebox.showerror(message='Cannot create shortcut', parent=self._tk_root, title=TITLE)
            return
        args = f'/c {program_arg} {dll_name}'

        path = filedialog.askdirectory(parent=self._tk_root,
                                       title='Choose location',
                                       mustexist=True)
        if not path:
            return
        path = os.path.abspath(path)

        shortcut = Dispatch('WScript.Shell').CreateShortCut(os.path.join(path, shortcut_name))
        shortcut.Targetpath = win_cmd_path
        shortcut.Arguments = args
        shortcut.WorkingDirectory = os.getcwd()
        shortcut.save()
        messagebox.showinfo(message=f'Shortcut {shortcut_name[:-4]} created', parent=self._tk_root, title=TITLE)

    def _on_destroy(self):
        self._switcher.save_data_json()
        self._tk_root.destroy()

    def mainloop(self):
        self._tk_root.mainloop()


def main():
    switcher = Switcher()

    parser = argparse.ArgumentParser(description='Switch to a different Ori dll')
    parser.add_argument('dll', nargs=argparse.OPTIONAL, help="the dll's name")
    args = parser.parse_args()
    if args.dll:
        dll_name = args.dll
        Tk().withdraw()  # Hides an extra window tk creates behind messagebox
        try:
            switcher.switch(dll_name)
            messagebox.showinfo(message=f'Switched to {dll_name} DLL', title=TITLE)
            sys.exit()
        except SwitcherException as e:
            messagebox.showerror(message=e, title=TITLE)
            sys.exit(1)

    ui = Ui(switcher)
    ui.mainloop()


if __name__ == '__main__':
    main()
