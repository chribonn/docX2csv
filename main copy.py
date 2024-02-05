import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter.filedialog import askopenfilename
import os.path
import tempfile


class UIScreen(tb.Frame):

    def __init__(self, master):
        super().__init__(master, padding=15)
        self.filename = tb.StringVar()
        self.racivalues = []
        self.pack(fill=BOTH, expand=YES)
        self.create_widget_elements()

        style = tb.Style()
        file_entry = tb.Entry(self, textvariable=self.filename, state=READONLY)
        file_entry.grid(row=0, column=0, columnspan=3, padx=20, pady=20, sticky="we")

        browse_btn = tb.Button(self, text="Browse", command=self.open_file)
        browse_btn.grid(row=0, column=1, padx=20, pady=20, sticky="e")
        
        raci_label = tb.Label(self, text="RACI Styles")
        raci_label.grid(row=1, column=0, padx=20, pady=20, sticky="w")

        raci_combo = tb.Combobox(self, values=self.racivalues)
        raci_combo.grid(row=1, column=1, columnspan=3, padx=20, pady=20, sticky="e")

    def open_file(self):
        # function populates this var
        xml_styles = ["A", "B", "C", "D"]
         
        # populate the raci_combo
        # self.master.raci_combo['values'] = xml_styles  # AttributeError: '_tkinter.tkapp' object has no attribute 'raci_combo'
        # self.master.racivalues = xml_styles  # did not work

        # self.raci_combo['values'] = xml_styles  # AttributeError: 'UIScreen' object has no attribute 'raci_combo'
        self.racivalues = xml_styles   # did not work

        # populate the text box
        self.filename.set("Done")


if __name__ == '__main__':

    app = tb.Window("combo", "sandstone", size=(800,400), resizable=(True, True))
    UIScreen(app)
    app.mainloop()
