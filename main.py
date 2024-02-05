import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter.filedialog import askopenfilename
import os.path
import tempfile

import docX2csv


class UIScreen(ttk.Frame):

    def __init__(self, master):
        super().__init__(master, padding=15)
        self.filename = ttk.StringVar()
        self.outputfl = ttk.StringVar()
        self.racivalues = []
        self.textvalues = []
        self.pack(fill=BOTH, expand=YES)
        self.create_widget_elements()

    def create_widget_elements(self):
        style = ttk.Style()
        
        #row 0
        file_entry = ttk.Entry(self, textvariable=self.filename, state=READONLY)
        file_entry.grid(row=0, column=0, columnspan=3, padx=20, pady=20, sticky="we")

        browse_btn = ttk.Button(self, text="Browse", command=self.open_file)
        browse_btn.grid(row=0, column=4, padx=20, pady=20, sticky="e")
        
        #row 1
        # The Styles that will be searched within the document
        raci_label = ttk.Label(self, text="RACI Styles")
        raci_label.grid(row=1, column=0, padx=20, pady=20, sticky="w")

        self.raci_combo = ttk.Combobox(self, values=self.racivalues, state=READONLY)
        self.raci_combo.grid(row=1, column=1, columnspan=3, padx=30, pady=20, sticky="we")

        #row 2
        # Optional Style that will be incorporated into the text
        text_label = ttk.Label(self, text="Style as text (Optional)")
        text_label.grid(row=2, column=0, padx=20, pady=20, sticky="w")

        self.text_combo = ttk.Combobox(self, values=self.textvalues, state=READONLY)
        self.text_combo.grid(row=2, column=1, padx=30, pady=20, sticky="we")

        #row 3
        # Location of the output file
        outputfl_label = ttk.Label(self, text="Output Filename")
        outputfl_label.grid(row=3, column=0, padx=20, pady=20, sticky="w")

        # Location of the output file
        outputfl_entry = ttk.Entry(self, textvariable=self.outputfl, state=READONLY)
        outputfl_entry.grid(row=3, column=1, padx=20, pady=20, sticky="w")

        #row 4
        browse_btn = ttk.Button(self, text="Run", command=self.process_file, state=DISABLED)
        browse_btn.grid(row=4, column=1, padx=20, pady=20, sticky="we")


    def process_file(self):
        pass


    def open_file(self):
        docX_file = askopenfilename()
        if not docX_file:
            return
        elif not os.path.isfile(docX_file):
            return
        elif not docX_file.endswith('.docx'):
            return
        
        tmp_dir = tempfile.TemporaryDirectory()
        xml_doc = os.path.join(docX2csv.extract_document_xml(docX_file, tmp_dir.name),docX2csv.XML_DOC_PATH.replace('/', '\\'))
        
        # extract the available styles
        xml_styles = docX2csv.get_styles(xml_doc)
        
        # populate the raci_combo
        self.raci_combo['values'] = xml_styles
        self.racivalues = xml_styles
        
        # populate the text_combo
        self.text_combo['values'] = xml_styles
        self.textvalues = xml_styles

        # populate the text box
        self.filename.set(docX_file)
        
        # generate the path and name of the csv files
        self.outputfl.set(os.path.join(os.path.dirname(docX_file), os.path.splitext(os.path.basename(docX_file))[0], '.csv'))


if __name__ == '__main__':
    app = ttk.Window("docX2csv", "sandstone", size=(800,400), resizable=(True, True))
    UIScreen(app)
    app.mainloop()
