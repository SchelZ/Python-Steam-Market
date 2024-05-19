## librarii necesare: openpyxl, bs4, tkinter, urllib.request

import urllib.request, pandas as pd
from openpyxl import load_workbook

from bs4 import BeautifulSoup
from typing import Final
from tkinter import Button, Entry, Frame, Tk, Toplevel, Canvas, END, DISABLED, CENTER, TOP, BOTH
from tkinter.filedialog import asksaveasfile

OA_database: list = None
OA_database_aux: list = None

class GraphWindow(Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("")
        self.wm_resizable(False, False)

        OA_height = 500

        OA_canva = Canvas(self, width=400, height=OA_height, bg='white')
        OA_canva.pack()

        OA_BAR_WIDTH = 55
        OA_BAR_SPACING = 20
        OA_CHART_TOP_MARGIN = 30

        OA_max_value = max(OA_database, key=lambda item: item[1])[1]
        OA_scale = (OA_height - OA_CHART_TOP_MARGIN) / OA_max_value

        OA_x = OA_BAR_SPACING
        for category, value in OA_database:
            if value == 0: bar_height = 10 * OA_scale
            else: bar_height = value * OA_scale
            OA_canva.create_text(OA_x + OA_BAR_WIDTH // 2, OA_height-10, text=category )
            OA_canva.create_rectangle(OA_x, OA_height - bar_height, OA_x + OA_BAR_WIDTH, OA_height - 25, fill="blue", outline="black")
            if value == 0: OA_canva.create_text(OA_x + OA_BAR_WIDTH // 2, OA_height - bar_height - 10, text="Free")
            else: OA_canva.create_text(OA_x + OA_BAR_WIDTH // 2, OA_height - bar_height - 10, text=f"{round(value, 1)/10} €")
            OA_x += OA_BAR_WIDTH + OA_BAR_SPACING

class DisplayMatrix(Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("")
        self.wm_resizable(False, False)
        
        OA_total_rows = len(OA_database_aux) 
        OA_total_columns = len(OA_database_aux[0]) 

        OA_max_value = max(OA_database_aux, key=lambda item: len(item[0]))
        for i in range(OA_total_rows): 
            for j in range(OA_total_columns): 
                self.OA_e = Entry(self, width=len(OA_max_value[0]), fg='blue', font=('Arial',10,'bold')) 
                self.OA_e.grid(row=i, column=j)
                if isinstance(OA_database_aux[i][j], float | int): 
                    if int(OA_database_aux[i][j]) == 0: self.OA_e.insert(END, "Free")
                    else: self.OA_e.insert(END, f"{int(OA_database_aux[i][j]) / 10} €")
                else: self.OA_e.insert(END, OA_database_aux[i][j])
                self.OA_e.config(state=DISABLED)
                 
class StartPage(Tk):
    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        Tk.wm_title(self, "OA GUI application")
        Tk.resizable(self, False, False)

        self.OA_product_name = None
        self.geometry("500x400")

        Frame(self).pack(side=TOP, fill=BOTH, expand = True)
        Button(self, text="Retrieve data", command= lambda: self.__retrieve_data()).place(relx=0.5, rely=0.3, anchor=CENTER)         ## Retrieve data
        Button(self, text="Create the graph", command= lambda: self.__create_graph()).place(relx=0.5, rely=0.4, anchor=CENTER)        ## Create the graph
        Button(self, text="Display the matrix", command= lambda: self.__display_matrix()).place(relx=0.5, rely=0.5, anchor=CENTER)      ## Display the matrix
        Button(self, text="Save to Excel file", command= lambda: self.__save_file()).place(relx=0.5, rely=0.6, anchor=CENTER)           ## Save to Excel file

    def __retrieve_data(self):
        global OA_database, OA_database_aux
        OA_website: Final[str] = "https://store.steampowered.com/search/?sort_by=Released_DESC&os=win"
        OA_headers: Final[dict] = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1941.0 Safari/537.36'}

        OA_req = urllib.request.Request(url=OA_website, headers=OA_headers)
        OA_webpage = urllib.request.urlopen(OA_req).read()
        OA_soup = BeautifulSoup(OA_webpage, features="html.parser")
        self.OA_product_name = [item.text.strip() for item in OA_soup.findAll("div", {"class": "ellipsis"})]
        OA_product_price: Final[list] = [item.text for item in OA_soup.findAll("div", {"class": "discount_final_price"})]
        for i, item in enumerate(OA_product_price):
            if item != "Free": OA_product_price[i] = float(item.replace('€', '').replace(",", "."))*10
            else: OA_product_price[i] = 0
        OA_database_aux = list(zip(self.OA_product_name, OA_product_price))
        OA_database = [item for item in OA_database_aux if len(item[0]) < 10]

    def __create_graph(self):
        if OA_database is not None:
            OA_graph_window = GraphWindow()

    def __display_matrix(self):
        if OA_database is not None:
            OA_graph_window = DisplayMatrix()

    def __save_file(self):
        if OA_database is not None:
            OA_Pret = {}
            OA_Nume = {}

            for i, (name, value) in enumerate(OA_database_aux):
                OA_Nume[i] = name
                if value == 0: OA_Pret[i] = 'Free'
                else:  OA_Pret[i] = f"{int(value) / 10} €"

            OA_result = {'Pret': OA_Pret, 'Nume': OA_Nume}
            OA_marks_data = pd.DataFrame(OA_result, columns=["Pret", "Nume"])
            OA_sytle = OA_marks_data.style.set_properties(**{'text-align': 'center'})
        
            OA_file = asksaveasfile(mode="w", initialfile ='Untitled.xlsx', defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
            
            OA_sytle.to_excel(OA_file.name, freeze_panes=(1,1))

            OA_workbook = load_workbook(OA_file.name)
            OA_worksheet = OA_workbook.active
            
            OA_worksheet.column_dimensions["B"].width = 13
            OA_worksheet.column_dimensions["C"].width = len(max(OA_database_aux, key=lambda item: len(item[0]))[0]) + 5

            OA_workbook.save(OA_file.name)

if __name__=="__main__":
    OA_app = StartPage()
    OA_app.mainloop()