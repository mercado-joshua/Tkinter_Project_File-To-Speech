#===========================
# Imports
#===========================

import tkinter as tk
from tkinter import ttk, colorchooser, Menu, Spinbox, scrolledtext, messagebox as mb, filedialog as fd

import os
from win32com.client import Dispatch

#===========================
# Main App
#===========================

class App(tk.Tk):
    """Main Application."""

    list_of_files = os.listdir()
    s = Dispatch('SAPI.SpVoice')

    #===========================================
    def __init__(self, title, icon, theme):
        super().__init__()
        self.style = ttk.Style(self)
        self.resizable(False, False)
        self.title(title)
        self.iconbitmap(icon)
        self.style.theme_use(theme)

        self.init_UI()
        self.init_events()

    # INITIALIZER ==============================
    @classmethod
    def create_app(cls, app):
        return cls(app['title'], app['icon'], app['theme'])

    #===========================================
    def init_events(self):
        self.bind('<<ListboxSelect>>', self.evt_show_content)
        self.listbox.bind('<Double-Button-1>', self.evt_open_file)

    #===========================================
    def init_UI(self):
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # ------------------------------------------
        self.listbox = tk.Listbox(self.main_frame)
        for file in self.list_of_files:
            if file.endswith(".txt"):
                self.listbox.insert(tk.END, file)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH)

        self.text = tk.Text(self.main_frame)
        self.text.pack(side=tk.LEFT)

        button = ttk.Button(self.main_frame, text='audio', command=self.audio)
        button.pack(side=tk.LEFT, anchor=tk.NW)

    # ------------------------------------------
    def audio(self):
        self.s.Speak(self.text.get('1.0', tk.INSERT))


    # EVENTS ------------------------------------
    def evt_show_content(self, event):
        x = self.listbox.curselection()[0]
        file = self.listbox.get(x)
        with open(file) as file:
            file = file.read()
        self.text.delete('1.0', tk.END)
        self.text.insert(tk.END, file)

    def evt_open_file(self, event):
        x = self.listbox.curselection()[0]
        os.system(self.listbox.get(x))

#===========================
# Start GUI
#===========================

def main(config):
    app = App.create_app(config)
    app.mainloop()

if __name__ == '__main__':
    main({
        'title' : 'Test',
        'icon' : 'python.ico',
        'theme' : 'clam'
        })