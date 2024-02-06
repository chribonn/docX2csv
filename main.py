import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import ttk
from ttkbootstrap.scrolled import ScrolledFrame
from tkinter.filedialog import askopenfilename
from ttkbootstrap.toast import ToastNotification
import xml.etree.ElementTree as ET
import os.path
import tempfile

import docX2csv


class UIScreen(tb.Frame):

    def __init__(self, master):
        super().__init__(master, padding=15)
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.filename = tb.StringVar()
        self.xmlfile = ''
        self.csv_fl = tb.StringVar()
        self.hdr_reset_on = tb.StringVar()
        self.cnt_style_tags = 0
        self.header_style = tb.StringVar()
        self.toast = ToastNotification(
            title='doc2Xcsv',
            message='Processing completed',
            duration=3000,
        )
        self.pack(fill=BOTH, expand=YES)
        self.create_widget_elements()
        

    def create_widget_elements(self):
        style = tb.Style()
        
        #row 0
        file_entry = tb.Entry(self, textvariable=self.filename, font=('Arial', 10), state=READONLY, width="60")
        file_entry.grid(row=0, column=0, columnspan=3, padx=20, pady=20, sticky="w")

        browse_btn = tb.Button(self, text="Browse", command=self.open_file)
        browse_btn.grid(row=0, column=4, padx=10, pady=20, sticky="w")
        
        #row 1
        # The Styles that will be searched within the document
        crossref_label = tb.Label(self, text="CrossRef Styles")
        crossref_label.grid(row=1, column=0, columnspan=1, padx=20, pady=20, sticky="w")

        self.crossref_scrollbar = tb.Scrollbar(self)
        self.crossref_treev = tb.Treeview(self, columns=('Style'), show='', yscrollcommand=self.crossref_scrollbar.set)
        self.crossref_treev.grid(row=1, column=1, columnspan=3, padx=10, pady=20, sticky="w")

        #row 2
        # Optional Style that will be incorporated into the text
        text_label = tb.Label(self, text="Header Style")
        text_label.grid(row=2, column=0, padx=10, pady=20, sticky="w")

        self.hdr_styl_combo = tb.Combobox(self, textvariable=self.header_style, state=READONLY)
        self.hdr_styl_combo.grid(row=2, column=1, padx=10, pady=20, sticky="w")

        # Optional Style that will be incorporated into the text
        text_label2 = tb.Label(self, text="Reset on")
        text_label2.grid(row=2, column=2, padx=20, pady=20, sticky="w")

        self.hdr_reset_combo = tb.Combobox(self, textvariable=self.hdr_reset_on, values=('Section', 'Page'), state=READONLY)
        self.hdr_reset_combo.current(0)
        self.hdr_reset_combo.grid(row=2, column=3, padx=10, pady=20, sticky="w")

        #row 3
        # Location of the output file
        csv_fl_label = tb.Label(self, text="Output Filename")
        csv_fl_label.grid(row=3, column=0, padx=20, pady=20, sticky="w")

        # Location of the output file
        csv_fl_entry = tb.Entry(self, textvariable=self.csv_fl, font=('Arial', 10), state=READONLY, width="60")
        csv_fl_entry.grid(row=3, column=1, columnspan=3, padx=20, pady=20, sticky="w")

        #row 4
        self.browse_btn = tb.Button(self, text="Run", command=self.process_file, state=DISABLED, width="60")
        self.browse_btn.grid(row=4, column=0, columnspan=3, padx=10, pady=20, sticky="we")


    def open_file(self):
        docX_file = askopenfilename()
        if not docX_file:
            return
        elif not os.path.isfile(docX_file):
            return
        elif not docX_file.endswith('.docx'):
            return
        
        self.xmlfile = os.path.join(docX2csv.extract_document_xml(docX_file, self.tmp_dir.name),docX2csv.XML_DOC_PATH.replace('/', '\\'))
        
        # extract the available styles
        xml_styles, self.cnt_style_tags = docX2csv.get_styles(self.xmlfile)
        
        # populate the crossref_treev
        cnt = 0
        for style in xml_styles:
            self.crossref_treev.insert('', tb.END, values=(style))
        
        # populate the text_combo
        self.hdr_styl_combo['values'] = xml_styles

        # populate the text box
        self.filename.set(docX_file)
        
        # generate the path and name of the csv files. This is identical to the source document except for a different extension
        self.csv_fl.set(os.path.splitext(self.filename.get())[0] + '.csv')
        
        self.browse_btn.config(state=NORMAL)
        

    def update_progressbar(self, style_cnt):
        pass


    def process_file(self):
        # styles that must be searched need to be defined
        if not self.crossref_treev.selection():
            return

        crossref_items = []
        for select in self.crossref_treev.selection():
            style = self.crossref_treev.item(select).get('values')[0]
            crossref_items.append(style)
            
        # Process the file
        parser = ET.XMLParser(encoding="utf-8")
        tree = ET.parse(self.xmlfile, parser=parser)

        root = tree.getroot()
        page = section = line = 1
        style_cnt = 0
        header_style = () + (self.header_style.get() ,)
        header_style_text = ''
        csvList = []


        ET.register_namespace("w", docX2csv.NS_URI)
        ns = {"w": docX2csv.NS_URI}
        # ET.dump(tree)

        for x in root.findall('.//w:p', ns):
            # print (x)
            style_text = ''
            style = None
            for y in x:
                if y.tag == docX2csv.NW_URI_TAG + 'pPr':
                    # The section check need to be on top because the node may not have a style
                    if docX2csv.proc_pPr_sectPr(y):
                        section += 1
                        if self.hdr_reset_on.get() == 'Section':
                            header_style_text = ''
                    
                    # If this is a Header Style is defined then extract text related to it
                    if self.header_style.get() != '':
                        style, styletag_found = docX2csv.proc_pPr_pStyle(y, header_style) or (None, False)
                        if style is not None:
                            if header_style_text != '':
                                header_style_text += '\n' 
                            header_style_text += docX2csv.proc_r_t(x)

                    # Process Cross Reference Styles
                    style, styletag_found = docX2csv.proc_pPr_pStyle(y, crossref_items) or (None, False)
                    if styletag_found:
                        style_cnt += 1
                        # update the progressbar
                        self.update_progressbar(style_cnt)
                    if style is None:
                        break
                    else:
                        style_text += docX2csv.proc_r_t(x)
                elif y.tag == docX2csv.NW_URI_TAG + 'r':
                    # Search for a rendered page breaks
                    if docX2csv.proc_r_lastRenderedPageBreak(y):
                        page += 1
                        line = 1
                        if self.hdr_reset_on.get() == 'Page':
                            header_style_text = ''
                    


            if style is not None:
                csvList.append(    
                    {
                        'Style' : style,
                        'StyleText' : style_text,
                        'HeaderStyleText' : header_style_text,
                        'Section' : section,
                        'Page': page,
                        'Line': line
                    })

                
            line += 1
        
        csv_flname = self.csv_fl.get()
        docX2csv.savecsv (csv_flname, csvList)
        self.toast.show_toast()


if __name__ == '__main__':
    app = tb.Window(f'docX2csv - {docX2csv.VERSION}', "sandstone", size=(800,640), resizable=(True, True))
    sf = ScrolledFrame(app, autohide=True)
    sf.pack(fill=BOTH, expand=YES, padx=10, pady=10)

    UIScreen(sf)
    app.mainloop()
