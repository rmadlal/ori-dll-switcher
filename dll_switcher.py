import os
import json
import shutil
import argparse
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
from win32com.client import Dispatch

TITLE = 'Ori DLL Switcher'
ORI_ROOT = r'C:\Program Files (x86)\Steam\steamapps\common\Ori DE'
ASSEMBLY_CSHARP = os.path.join(ORI_ROOT, 'oriDE_Data', 'Managed', 'Assembly-CSharp.dll')
DLL_NAMES = 'dll_names.json'


class Switcher(object):
    def __init__(self):
        self.dll_names = self.load_dll_names()
        self.list = self.dll_names.keys()
        self.dll_path = ''
        self.dll_label_hint = 'Please locate your specified DLL.'

        self.tk_root = Tk()
        self.tk_root.title(TITLE)
        self.tk_root.columnconfigure(0, weight=1)
        self.tk_root.rowconfigure(2, weight=1)
        self.tk_root.geometry('350x120')
        self.tk_root.protocol('WM_DELETE_WINDOW', self._on_destroy)

        self.combobox = Combobox(self.tk_root, values=tuple(self.dll_names.keys()), state="readonly")
        self.combobox.grid(row=0, column=0, padx=8, pady=8, sticky=W+E)
        self.combobox.bind('<<ComboboxSelected>>', lambda evt: self._update_chosen_dll())
        self.combobox.bind('<Return>', lambda evt: self._add_new_dll_name(self.combobox.get()))
        if self.dll_names:
            self.combobox.current(0)

        self.button_new = Button(self.tk_root, text='New', command=self._button_new)
        self.button_new.grid(row=0, column=1, padx=8, sticky=W+E)

        self.stringvar_label_dll = StringVar()
        self.label_dll = Label(self.tk_root, textvariable=self.stringvar_label_dll)
        self.label_dll.grid(row=1, column=0, padx=8, sticky=W)

        self.button_locate = Button(self.tk_root, text='Locate DLL', command=self._locate_dll)
        self.button_locate.grid(row=1, column=1, padx=8, sticky=W+E)

        self.button_apply = Button(self.tk_root, text='Apply', command=self._apply, state=DISABLED)
        self.button_apply.grid(row=2, column=0, padx=8, pady=16, sticky=S)

        self.button_create_shortcut = Button(
            self.tk_root, text='Create shortcut', command=self._create_shortcut, state=DISABLED)
        self.button_create_shortcut.grid(row=2, column=1, padx=8, pady=16, sticky=W+E+S)

        self.ori_root, self.assembly_csharp = self._validate_ori_root()
        self._update_chosen_dll()

    @staticmethod
    def load_dll_names():
        try:
            with open(DLL_NAMES) as f:
                return json.load(f)
        except OSError:
            return {}

    def _store_dll_names(self):
        with open(DLL_NAMES, 'w') as f:
            json.dump(self.dll_names, f)

    def _validate_ori_root(self):
        ori_root = ORI_ROOT
        assembly_csharp = ASSEMBLY_CSHARP
        while not os.path.exists(os.path.dirname(assembly_csharp)):
            path = filedialog.askdirectory(parent=self.tk_root,
                                           title='Select the Ori DE directory',
                                           mustexist=True)
            if not path:
                messagebox.showerror(message='Ori DE directory not found', parent=self.tk_root, title=TITLE)
                sys.exit(1)

            ori_root = os.path.abspath(path)
            assembly_csharp = os.path.join(ori_root, 'oriDE_Data', 'Managed', 'Assembly-CSharp.dll')
            if os.path.exists(os.path.dirname(assembly_csharp)):
                break

            messagebox.showerror(
                message='Invalid Ori DE directory. Should be the directory that contains oriDE.exe.',
                parent=self.tk_root, title=TITLE)

        return ori_root, assembly_csharp

    # new dll path was set
    def _validate_path(self):
        dll_name = self.combobox.get()
        if self.dll_path:
            self.stringvar_label_dll.set(self.dll_path.split(os.sep)[-1])
            self.label_dll.configure(state=NORMAL)
            self.button_apply.configure(state=NORMAL)
            self.button_create_shortcut.configure(state=NORMAL)
        else:
            self.stringvar_label_dll.set(self.dll_label_hint if dll_name else '')
            self.label_dll.configure(state=DISABLED)
            self.button_apply.configure(state=DISABLED)
            self.button_create_shortcut.configure(state=DISABLED)

    # user chose a name from the list
    def _update_chosen_dll(self):
        dll_name = self.combobox.get()
        self.dll_path = self.dll_names[dll_name] if dll_name in self.dll_names else ''
        self._validate_path()
        self.button_locate.configure(state=NORMAL if dll_name else DISABLED)

    # user clicked on button_new
    def _button_new(self):
        self.combobox.configure(state=NORMAL)
        self.combobox.set('')
        self.combobox.focus()
        self._update_chosen_dll()

    # user pressed enter after typing a new name
    def _add_new_dll_name(self, name):
        if not name:
            return
        self.dll_names[name] = ''
        self.combobox.configure(values=tuple(self.dll_names.keys()), state="readonly")
        self._update_chosen_dll()

    # user clicked on button_locate
    def _locate_dll(self):
        dll_name = self.combobox.get()
        path = filedialog.askopenfilename(filetypes=[('DLL', '*.dll')],
                                          initialdir=self.ori_root,
                                          parent=self.tk_root,
                                          title=f'Choose your "{dll_name}" DLL file')
        if path:
            self.dll_path = os.path.abspath(path)
            self.dll_names[dll_name] = self.dll_path
            self._validate_path()

    # user clicked on button_apply
    def _apply(self):
        shutil.copy(self.dll_path, self.assembly_csharp)
        messagebox.showinfo(message='Done!', parent=self.tk_root, title=TITLE)
        self._on_destroy()

    # user clicked on button_create_shortcut
    def _create_shortcut(self):
        dll_name = self.combobox.get()
        shortcut_name = f'SwitchTo{dll_name.capitalize()}.lnk'
        win_cmd_path = r'C:\Windows\System32\cmd.exe'
        program_arg = os.path.basename(sys.argv[0])

        if program_arg.endswith('.py'):
            program_arg = 'python ' + program_arg
        elif program_arg.endswith('.exe'):
            program_arg = program_arg[:-4]
        else:
            messagebox.showerror(message='Cannot create shortcut', parent=self.tk_root, title=TITLE)
            return
        args = f'/c {program_arg} {dll_name}'

        path = filedialog.askdirectory(parent=self.tk_root,
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
        messagebox.showinfo(message='Shortcut created', parent=self.tk_root, title=TITLE)

    def _on_destroy(self):
        self._store_dll_names()
        self.tk_root.destroy()


def main():
    parser = argparse.ArgumentParser(description='Switch to a different Ori dll.')
    parser.add_argument('dll', nargs=argparse.OPTIONAL, help="the dll's name")
    args = parser.parse_args()
    if args.dll:
        dll_names = Switcher.load_dll_names()
        if args.dll not in dll_names:
            sys.exit(f'"{args.dll}" DLL not found')
        shutil.copy(dll_names[args.dll], ASSEMBLY_CSHARP)
        print('Done!')
        return

    gui = Switcher()
    gui.tk_root.mainloop()


if __name__ == '__main__':
    main()
