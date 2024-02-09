import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import ttk
from ttkbootstrap.scrolled import ScrolledFrame
from tkinter.filedialog import askopenfilename
from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.dialogs import Messagebox
import xml.etree.ElementTree as ET
import os.path
import tempfile
import csv
import uuid

import docX2csv_lib


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
        docXfile_label = tb.Label(self, text="DocX File", width="15")
        docXfile_label.grid(row=0, column=0, columnspan=1, padx=5, pady=20, sticky="w")

        file_entry = tb.Entry(self, textvariable=self.filename, font=('Arial', 10), state=READONLY, width="40")
        file_entry.grid(row=0, column=1, columnspan=2, padx=1, pady=20, sticky="w")

        browse_btn = tb.Button(self, text="Browse", command=self.open_file, width="15")
        browse_btn.grid(row=0, column=3, columnspan=1, padx=0, pady=20, sticky="e")
        
        #row 1
        # The Styles that will be searched within the document
        crossref_label = tb.Label(self, text="CrossRef Styles", width="15")
        crossref_label.grid(row=1, column=0, columnspan=1, padx=5, pady=20, sticky="w")

        self.crossref_scrollbar = tb.Scrollbar(self)
        self.crossref_treev = tb.Treeview(self, columns=('Style'), show='', yscrollcommand=self.crossref_scrollbar.set)
        self.crossref_treev.grid(row=1, column=1, columnspan=2, padx=1, pady=20, sticky="w")

        #row 2
        # Optional Style that will be incorporated into the text
        text_label = tb.Label(self, text="Header Style", width="15")
        text_label.grid(row=2, column=0, columnspan=1, padx=5, pady=20, sticky="w")

        self.hdr_styl_combo = tb.Combobox(self, textvariable=self.header_style, state=READONLY, width="20")
        self.hdr_styl_combo.grid(row=2, column=1, columnspan=1, padx=1, pady=20, sticky="w")

        # Optional Style that will be incorporated into the text
        text_label2 = tb.Label(self, text="Align on", width="10")
        text_label2.grid(row=2, column=2, columnspan=1, padx=20, pady=20, sticky="e")

        self.hdr_reset_combo = tb.Combobox(self, textvariable=self.hdr_reset_on, values=('Section', 'Page'), state=READONLY, width="15")
        self.hdr_reset_combo.current(0)
        self.hdr_reset_combo.grid(row=2, column=3, padx=0, pady=20, sticky="w")

        #row 3
        # Location of the output file
        csv_fl_label = tb.Label(self, text="Output Filename", width="15")
        csv_fl_label.grid(row=3, column=0, columnspan=1, padx=5, pady=20, sticky="w")

        # Location of the output file
        csv_fl_entry = tb.Entry(self, textvariable=self.csv_fl, font=('Arial', 10), state=READONLY, width="55")
        csv_fl_entry.grid(row=3, column=1, columnspan=3, padx=1, pady=20, sticky="w")

        #row 4
        blank2_label = tb.Label(self, text=" ", width="15")
        blank2_label.grid(row=4, column=1, columnspan=1, padx=5, pady=20, sticky="w")

        self.browse_btn = tb.Button(self, text="Run", command=self.process_file, state=DISABLED, width="40")
        self.browse_btn.grid(row=4, column=1, columnspan=2, padx=0, pady=20, sticky="we")
        
        #row 5 - Add Progress bar


    def open_file(self):
        docX_file = askopenfilename(filetypes=[('Word files','*.docx')])
        if not docX_file:
            return
        elif not os.path.isfile(docX_file):
            return
        elif not docX_file.endswith('.docx'):
            return
        
        self.xmlfile = os.path.join(docX2csv_lib.extract_document_xml(docX_file, self.tmp_dir.name),docX2csv_lib.XML_DOC_PATH.replace('/', '\\'))
        
        # extract the available styles
        xml_styles, self.cnt_style_tags = docX2csv_lib.get_styles(self.xmlfile)
        
        # populate the crossref_treev
        cnt = 0
        for style in xml_styles:
            self.crossref_treev.insert('', tb.END, values=(style))
        
        # populate the text_combo
        xml_styles.insert(0, '')
        self.hdr_styl_combo['values'] = xml_styles

        # populate the text box
        self.filename.set(docX_file)
        
        # generate the path and name of the csv files. This is identical to the source document except for a different extension
        self.csv_fl.set(os.path.splitext(self.filename.get())[0] + '.csv')
        
        self.browse_btn.config(state=NORMAL)
        

    def update_progressbar(self, style_cnt):
        pass


    def process_file(self):
        def updcsv(csvList, style, style_text, header_style_text, section, page, line, uuid):
            csvList.append(
            {
                'Style' : style,
                'Style Text' : style_text,
                'Header Style Text' : header_style_text,
                'Section' : section,
                'Page': page,
                'Line': line,
                'Linked Ref': uuid
            })
        
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
        header_style_dict = {}
        crossref_style_dict = {}
        
        # Because of a quirk in the docx xml format there can be two page breaks on two adjacent
        # './/w:p' nodes one related to a style and the other being a <w:lastRenderedPageBreak/>
        # in this scenario only  count as one page.
        pagebreak_prior = False

        ET.register_namespace("w", docX2csv_lib.NS_URI)
        ns = {"w": docX2csv_lib.NS_URI}
        # ET.dump(tree)

        for x in root.findall('.//w:p', ns):
            # print (x)
            style_text = ''
            style = None
            if docX2csv_lib.page_break(x):
                if not pagebreak_prior:
                    page += 1
                    line = 1
                pagebreak_prior = True
            else:
                pagebreak_prior = False
                
                    
            
            for y in x:
                if y.tag == docX2csv_lib.NW_URI_TAG + 'pPr':
                    # The section check need to be on top because the node may not have a style
                    if docX2csv_lib.proc_pPr_sectPr(y):
                        section += 1

                    # If this is a Header Style extract text related and file it in the Dictionary
                    if self.header_style.get() != '':
                        style, styletag_found = docX2csv_lib.proc_pPr_pStyle(y, header_style) or (None, False)
                        if style is not None:
                            header_style_dict[uuid.uuid4().node] = (section if self.hdr_reset_on.get() == 'Section' else page, docX2csv_lib.proc_r_t(x))
                            break

                    # Process Cross Reference Styles
                    style, styletag_found = docX2csv_lib.proc_pPr_pStyle(y, crossref_items) or (None, False)
                    if styletag_found:
                        style_cnt += 1
                        # update the progressbar
                        self.update_progressbar(style_cnt)
                    if style is None:
                        break
                    else:
                        crossref_style_dict[uuid.uuid4().node] = (style, docX2csv_lib.proc_r_t(x), section, page, line)
                
            line += 1

        csvList = []
                    
        for x in crossref_style_dict:
            style = crossref_style_dict[x][0]
            style_text = crossref_style_dict[x][1]
            section = crossref_style_dict[x][2]
            page = crossref_style_dict[x][3]
            line = crossref_style_dict[x][4]
            header_style_text = uuidref = ''
            
            # if a match is found with the header data generate a uuid and don't process an empty header
            if self.header_style.get() != '':
                for y in header_style_dict:
                    if header_style_dict[y][0] == (section if self.hdr_reset_on.get() == 'Section' else page):
                        header_style_text = header_style_dict[y][1]
                        uuidref = f'{style}{uuid.uuid4().node}' if uuidref == '' else uuidref
                        updcsv(csvList, style, style_text, header_style_text, section, page, line, uuidref)
                # if no header record was found still write the style record
                if header_style_text == '':
                    updcsv(csvList, style, style_text, header_style_text, section, page, line, uuidref)
            else:
                header_style_text = ''
                updcsv(csvList, style, style_text, header_style_text, section, page, line, uuidref)

        csv_flname = self.csv_fl.get()
        csvColumns = ['Style','Style Text','Header Style Text','Section','Page','Line','Linked Ref']
        try:
            with open(csv_flname, 'w', newline='') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=csvColumns)
                writer.writeheader()
                for data in csvList:
                    writer.writerow(data)
        except IOError:
            print("I/O error")

        self.toast.show_toast()
        app.destroy()


if __name__ == '__main__':
    app = tb.Window(f'docX2csv - {docX2csv_lib.VERSION}', "sandstone", size=(800,640), resizable=(True, True))
    try:
        app.iconbitmap('assets/docX2csv.ico')
    except:
        # look for the file with the program
        try:
            app.iconbitmap('docX2csv.ico')
        except:
            app.iconbitmap('')
            Messagebox.show_error(title = 'Icon file: docX2csv.ico not found!', message='It should be in the ASSETS folder\nlocated under this exe.')
    sf = ScrolledFrame(app, autohide=True)
    sf.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    UIScreen(sf)
    app.mainloop()

