from os import path
from tkinter import Label, Entry, Button, Menu, StringVar, filedialog, Tk, END
from tkinter import messagebox as mb
import tkinter
import webbrowser

from numpy import pad

import settings

class App(Tk):
    def __init__(self):
        super().__init__()

        ## App Settings
        self.resizable(False, False)
        path_to_icon = path.abspath(path.join(path.dirname(__file__), 'icon.ico'))
        self.iconbitmap(path_to_icon)
        self.title('Date Verification Utility')
        self.version = "0.1.0"
        self.author = "Paul Runco"

        self.settings = settings.load()

        ## Menu
        self.menubar = Menu(self)

        self.filemenu = Menu(self.menubar, tearoff=False)
        self.filemenu.add_command(label="Exit", command=self.quit, accelerator="Ctrl+q")
        self.filemenu.add_command(label="Settings", command=lambda: self.display_settings(app, self.settings), accelerator="Ctrl+.")
        self.menubar.add_cascade(label="File", menu=self.filemenu)

        self.helpmenu = Menu(self.menubar, tearoff=False)
        self.helpmenu.add_command(label="About", command=self.open_about)
        self.helpmenu.add_command(label="Documentation", command=self.open_docs, accelerator="Ctrl+?")
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)

        self.bind_all("<Control-q>", exit)
        self.bind_all("<Control-?>", self.open_docs)
        self.bind_all("<Control-.>", self.display_settings)

        self.config(menu=self.menubar)

        ## Estimated Completion Dates Entry
        # Label
        self.estimated_completion_label = Label(
            self,
            text="Estimated Completion Dates",
            ).pack(padx=(5,0), anchor='w')
        # Entry
        self.estimated_completion_entry = Entry(
            self, 
            width=80, 
            textvariable=StringVar
            )
        self.estimated_completion_entry.pack(padx=5, pady=(0, 5))
        # Button
        self.estimated_completion_button = Button(
            self, text="Browse", 
            command=lambda: self.browse_for("estimated_completion")
            ).pack(padx=5, pady=(0, 5), anchor='e')

        ## Promise Date Verification Template Entry
        # Label
        self.date_verification_label = Label(
            self,
            text="Date Verification Template"
            ).pack(padx=(5,0), anchor='w')
        # Entry
        self.date_verification_entry = Entry(
            self,
            width=80,
            textvariable=StringVar
            )
        self.date_verification_entry.pack(padx=5, pady=(0, 5))
        # Button
        self.date_verification_button = Button(
            self, text="Browse", 
            command=lambda: self.browse_for("date_verification")
            ).pack(padx=5, pady=(0, 5), anchor='e')

        self.generate_button = Button(
            self, text="Update report",
            command=lambda: self.on_click_update()
        ).pack(fill='x', padx=5, pady=5)

    def browse_for(self, target):
        file_name = filedialog.askopenfilename(
            filetypes=(("Excel files", "*xlsx"), ("All files", "*"))
        )
        if target == "estimated_completion":
            self.estimated_completion_entry.config(background='white')
            self.estimated_completion_entry.delete(0, END)
            self.estimated_completion_entry.insert(0, file_name)
        if target == 'date_verification':
            self.date_verification_entry.config(background='white')
            self.date_verification_entry.delete(0, END)
            self.date_verification_entry.insert(0, file_name)

    def on_click_update(self):
        path_to_estimated_completion = self.estimated_completion_entry.get()
        if path_to_estimated_completion == "":
            mb.showwarning(title="Warning: ID-10T", message="Please select an Estimated Completion File")
            self.estimated_completion_entry.config(background= 'red')
            return

        path_to_date_verification = self.date_verification_entry.get()
        if path_to_date_verification == "":
            mb.showwarning(title="Warning: ID-10T", message="Please select a Date Verification File")
            self.date_verification_entry.config(background='red')
            return

    def display_settings(self, event=None):
        settings = self.settings
        settings_window = tkinter.Toplevel(self)
        settings_window.title('Settings')

        self.transit_time_label = Label(
            settings_window,
            text="Transit Time Days",
            ).grid(row=0, padx=(5,0))

        self.transit_time_entry = Entry(
            settings_window, 
            width=20, 
            textvariable=StringVar,
            )
        self.transit_time_entry.grid(row=0, column=1, padx=(0, 5))
        self.transit_time_entry.insert(0, settings['ShippingOptions']['TransitDays'])

    def open_about(self):
        mb.showinfo(
            title="About",
            message=f"Version {self.version} | {self.author}"
            )
    
    def open_docs(self, event=None):
        webbrowser.open('https://github.com/paulrunco/date-verification')


if __name__=="__main__":
    app = App()
    app.mainloop()