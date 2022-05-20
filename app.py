from os import path
from tkinter import BooleanVar, Label, Entry, Checkbutton, Button, Menu, StringVar, filedialog, Tk, END
from tkinter import messagebox as mb
import webbrowser

import settings
import functions

class App(Tk):
    def __init__(self):
        super().__init__()

        ## App Settings
        self.resizable(False, False)
        path_to_icon = path.abspath(path.join(path.dirname(__file__), 'icon.ico'))
        self.iconbitmap(path_to_icon)
        self.title('Date Verification Utility')
        self.version = "1.0.0"
        self.author = "Paul Runco"

        self.settings = settings.load()

        ## Menu
        self.menubar = Menu(self)

        self.filemenu = Menu(self.menubar, tearoff=False)
        # self.filemenu.add_command(label="Settings", command=lambda: self.display_settings(self), accelerator="Ctrl+.")
        self.filemenu.add_command(label="Clear", command=self.clear_form, accelerator="Ctrl+Delete")
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=self.quit, accelerator="Ctrl+Q")
        self.menubar.add_cascade(label="File", menu=self.filemenu)

        self.helpmenu = Menu(self.menubar, tearoff=False)
        self.helpmenu.add_command(label="About", command=self.open_about)
        self.helpmenu.add_command(label="Documentation", command=self.open_docs, accelerator="Ctrl+?")
        self.menubar.add_cascade(label="Help", menu=self.helpmenu)

        self.bind_all("<Control-q>", exit)
        self.bind_all("<Control-?>", self.open_docs)
        self.bind_all("<Control-Delete>", self.clear_form)

        self.config(menu=self.menubar)

        ## Estimated Completion Dates Entry
        # Label
        self.estimated_completion_label = Label(
            self,
            text="Estimated Completion Dates",
            ).grid(row=0, column=0, padx=(5,0), columnspan=2, sticky='w')
        # Entry
        self.estimated_completion_entry = Entry(
            self, 
            width=80, 
            textvariable=StringVar
            )
        self.estimated_completion_entry.grid(row=1, column=0, padx=5, pady=(0, 5), columnspan=2, sticky='w')
        # Button
        self.estimated_completion_button = Button(
            self, text="Browse", 
            command=lambda: self.browse_for("estimated_completion")
            ).grid(row=1, column=2, padx=5, pady=(0, 5), sticky='w')

        ## Promise Date Verification Template Entry
        # Label
        self.date_verification_label = Label(
            self,
            text="Date Verification Template"
            ).grid(row=2, column=0, padx=(5,0), columnspan=2, sticky='w')
        # Entry
        self.date_verification_entry = Entry(
            self,
            width=80,
            textvariable=StringVar
            )
        self.date_verification_entry.grid(row=3, column=0, padx=5, pady=(0, 5), columnspan=2, sticky='w')
        # Button
        self.date_verification_button = Button(
            self, 
            text="Browse", 
            command=lambda: self.browse_for("date_verification")
            ).grid(row=3, column=2, padx=5, pady=(0, 5), sticky='w')

        ## Settings
        self.transit_time_label = Label(
            self,
            text="Transit Time Days",
            ).grid(row=4, column=0, sticky='e')

        self.transit_days = StringVar(value=self.settings['ShippingOptions']['TransitDays'])
        self.transit_days.trace_add('write', self.save_settings)
        self.transit_time_entry = Entry(
            self,
            width=4, 
            textvariable=self.transit_days,
            justify='center'
            )
        self.transit_time_entry.grid(row=4, column=1, sticky='w')

        self.include_weekends = BooleanVar(value=self.settings['ShippingOptions']['IncludeWeekends'])
        self.include_weekends_checkbox = Checkbutton(
            self,
            text="Include Weekends",
            variable=self.include_weekends,
            onvalue=True, 
            offvalue=False,
            command=self.save_settings
        ).grid(row=4, column=1)

        self.update_button = Button(
            self, text="Update report",
            command=self.on_click_update
        )
        self.update_button.grid(row=5, padx=5, pady=5, columnspan=3, sticky='ew')

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

        self.reset_button()

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

        transit_days = int(self.transit_time_entry.get())
        include_weekends = self.include_weekends.get()
        if transit_days == "":
            mb.showwarning(title='Missing values', message='Please enter a number for Transit Days')
            self.transit_time_entry.config(background='red')
            return

        self.update_button.configure(text = 'Working...')
        functions.update_report(path_to_date_verification, path_to_estimated_completion, transit_days, include_weekends)
        self.update_button.configure(text = 'Done!')


    def save_settings(self, *args):
        transit_days = self.transit_days.get()
        include_weekends = self.include_weekends.get()
        if transit_days:
            self.settings['ShippingOptions'] = {'TransitDays': transit_days, 'IncludeWeekends': include_weekends}
        else:
            self.settings['ShippingOptions'] = {'TransitDays': 3, 'IncludeWeekends': include_weekends}
        settings.save(self.settings)
        self.reset_button()

    def reset_button(self):
        self.update_button.configure(text = 'Update report')

    def clear_form(self, *args):
        self.estimated_completion_entry.delete(0, END)
        self.date_verification_entry.delete(0, END)
        self.reset_button()

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