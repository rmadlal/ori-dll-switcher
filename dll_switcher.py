import argparse
import json
import os
import shutil
import sys
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox

from win32com.client import Dispatch

TITLE = 'Ori DLL Switcher'
ORI_ROOT = r'C:\Program Files (x86)\Steam\steamapps\common\Ori DE'
ASSEMBLY_CSHARP = r'oriDE_Data\Managed\Assembly-CSharp.dll'
SWITCHER_JSON_PATH = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), 'dll_switcher.json')
KEY_ROOT = 'root'
KEY_DLL_NAMES = 'dll_names'


class SwitcherException(Exception):
    pass


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
            with open(SWITCHER_JSON_PATH) as f:
                data = json.load(f)
                if not self._check_data_valid(data):
                    return None
                return data
        except OSError:
            return None

    def save_data_json(self):
        with open(SWITCHER_JSON_PATH, 'w') as f:
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

    def get_dll_path(self, name):
        return self.dll_names[name] if name in self.dll_names else None

    def set_dll_path(self, name, path):
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
        self._tk_root.columnconfigure(1, weight=1)
        self._tk_root.rowconfigure(2, weight=1)
        self._tk_root.geometry('560x120')
        self._tk_root.protocol('WM_DELETE_WINDOW', self._on_destroy)

        # UI elements
        # |---------|-----------------|
        # | Name:   | dropdown        |
        # | Set DLL | dll path text   |
        # | Apply   | Create shortcut |

        self._text_dll_name = Label(self._tk_root, text='Name:', anchor=W)
        self._text_dll_name.grid(row=0, column=0, padx=8, sticky=W)

        self._combobox = Combobox(self._tk_root, values=tuple(self._switcher.dll_names),
            validate='key', validatecommand=(self._tk_root.register(self._dll_name_typed), '%P'))
        self._combobox.grid(row=0, column=1, padx=8, pady=8, sticky=W)
        self._combobox.bind('<<ComboboxSelected>>', lambda _: self._dll_name_selected())
        if self._switcher.dll_names:
            self._combobox.current(0)

        self._button_set_dll = Button(self._tk_root, text='Set DLL', command=self._set_dll)
        self._button_set_dll.grid(row=1, column=0, padx=8, sticky=W)

        self._text_dll_path = Label(self._tk_root, anchor=W)
        self._text_dll_path.grid(row=1, column=1, columnspan=2, padx=8, sticky=W)

        self._button_apply = Button(self._tk_root, text='Apply', command=self._apply)
        self._button_apply.grid(row=2, column=0, columnspan=2, padx=8, pady=8, sticky=W + S)

        self._button_create_shortcut = Button(
            self._tk_root, text='Create shortcut', command=self._create_shortcut)
        self._button_create_shortcut.grid(row=2, column=1, padx=8, pady=8, sticky=W + S)

        self._validate_ori_root()
        self._dll_name_selected()

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
    def _update_dll_path(self, dll_path):
        self._text_dll_path.configure(text=dll_path or '')
        self._button_apply.configure(state=NORMAL if dll_path else DISABLED)
        self._button_create_shortcut.configure(state=NORMAL if dll_path else DISABLED)

    # dll name changed
    def _update_dll_name(self, dll_name):
        self._button_set_dll.configure(state=NORMAL if dll_name else DISABLED)
        dll_path = self._switcher.get_dll_path(dll_name)
        self._update_dll_path(dll_path)

    # user chose a name from the list
    def _dll_name_selected(self):
        self._update_dll_name(self._combobox.get())

    # user is typing a name (combobox validatecommand callback)
    def _dll_name_typed(self, text):
        self._update_dll_name(text)
        return True

    # user clicked on button_set_dll
    def _set_dll(self):
        dll_name = self._combobox.get()
        dll_path = filedialog.askopenfilename(filetypes=[('DLL', '*.dll')],
                                              initialdir=self._switcher.ori_root,
                                              parent=self._tk_root,
                                              title=f'Choose your "{dll_name}" DLL file')
        if not dll_path:
            return
        dll_path = os.path.abspath(dll_path)
        if dll_path == self._switcher.assembly_csharp:
            messagebox.showerror(message="You cannot choose the game's original DLL path.\nIf this is your mod DLL, please rename it or copy it elsewhere.",
                                 title=TITLE)
            return

        self._switcher.set_dll_path(dll_name, dll_path)
        self._update_dll_path(dll_path)
        self._combobox.configure(values=tuple(self._switcher.dll_names))

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
        program_dir, program_filename = os.path.split(os.path.abspath(sys.argv[0]))

        if program_filename.endswith('.py'):
            program_filename = 'py ' + program_filename
        elif not program_filename.endswith('.exe'):
            messagebox.showerror(message='Cannot create shortcut', parent=self._tk_root, title=TITLE)
            return
        args = f'/c {program_filename} {dll_name}'

        shortcut_dir = filedialog.askdirectory(parent=self._tk_root,
                                               title='Choose location',
                                               mustexist=True)
        if not shortcut_dir:
            return
        shortcut_dir = os.path.abspath(shortcut_dir)

        shortcut = Dispatch('WScript.Shell').CreateShortCut(os.path.join(shortcut_dir, shortcut_name))
        shortcut.Targetpath = win_cmd_path
        shortcut.Arguments = args
        shortcut.WorkingDirectory = program_dir
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
