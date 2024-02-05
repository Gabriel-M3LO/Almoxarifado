import openpyxl
import tkinter as tk
from tkinter import ttk
import pandas as pd
import customtkinter as ctk


#Pandas
tb = pd.read_excel('tabelas.xlsx')
TpMaterial = ["Escritório", "Penso", "Limpeza"]


#backend
def load_data():
    path = r"C:\Users\gabri\Documents\Almoxarifado\Almoxarifado\bd.xlsx"
    wb = openpyxl.load_workbook(path)
    sheet = wb.active

    list_values = list(sheet.values)
    print(list_values)

    for col_name in cols:
        treeview.heading(col_name, text=col_name)

    for values_tuple in list_values[0:]:
        treeview.insert('',tk.END, values = values_tuple[0:])
def insert_data():
    nome = material_entry.get()
    qnt = int(num_spinbox.get())
    Departamento = dep_pack.get()
    Tcovid = "Teste Covid" if a.get() else "Material"

    #save data in excel
    path = r"C:\Users\gabri\Documents\Almoxarifado\Almoxarifado\bd.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [nome, qnt, Departamento, Tcovid]
    sheet.append(row_values)
    workbook.save(path)

    #insert data save in previwer
    treeview.insert('', tk.END, values=row_values)

    #Clear form
    material_entry.delete(0, 'end')
    material_entry.insert(0, "name")

    num_spinbox.delete(0,'end')
    num_spinbox.insert(0, "num")

    dep_pack.set(deplist[0])

    checkbutton.state(["!selected"])

#frontend
#stating a window
root = ctk.CTk()
root.title("Almoxarifado")

def login ():
    root.geometry("400x500")

#configuring the theme
style = ttk.Style(root)
root.tk.call("source", "forest-dark.tcl")
root.tk.call("source", "forest-light.tcl")
style.theme_use("forest-dark")

#main frame
frame = ttk.Frame(root)
frame.pack(expand=True, fill='both')

widgets_frame = ttk.LabelFrame(frame, text="Insira os dados")
widgets_frame.grid(column=0, row=0, padx=10, pady=10, sticky="nsew")

#input informations
material_entry = ttk.Combobox(widgets_frame, values=TpMaterial)
material_entry.set("Tipo de material")
material_entry.grid(column=0, row=0, padx=5, pady=5, sticky="ew")

name_entry = ttk.Entry(widgets_frame, width=50)
name_entry.insert(0, "Produto...")
#name_entry.bind("<FocusIn>", lambda event:name_entry.delete(0, "end"))
name_entry.grid(column=1, row=0, padx=5, pady=5, sticky="ew", columnspan=2)

num_spinbox = ttk.Spinbox(widgets_frame, from_=0, to =10000, width=6)
num_spinbox.insert(0, "Qnt")
num_spinbox.grid(column=0, row=1, padx=5, pady=5, sticky="ew")

deplist = ["CERPAT","CEREST", "V. Epidemiológica", "V. Sanitária", "V.Ambiental", "Laboratório", "Zoonoses", "Endemias"]
dep_pack = ttk.Combobox(widgets_frame, values=deplist)
dep_pack.set("Departamentos...")
dep_pack.grid(column=1, row=1, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Teste Covid", variable=a)
checkbutton.grid(column=4, row=0, padx=5, pady=5,sticky="nsew")

buttoninput = ttk.Button(widgets_frame, text="Inserir Dados".upper(), command=insert_data)
buttoninput.grid(column=0, row=2, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid (column=0, row=3, padx=5, pady=10, sticky="nsew",columnspan=5)

#Creating previwer
tree_frame = ttk.Frame(frame)
tree_frame.grid(column=1, row=0, padx=(0,20), pady=(10))
tree_scrollbar = ttk.Scrollbar(tree_frame)
tree_scrollbar.pack(side="right", fill="y")

cols = ("Produto","Quantidade","Departamento","Valor unitário")
treeview = ttk.Treeview(tree_frame, show="headings", yscrollcommand=tree_scrollbar.set, columns = cols, height=13)

treeview.column("Produto", width=150)
treeview.column("Departamento", width=120)
treeview.column("Quantidade", width=100)
treeview.column("Valor unitário", width=100)
treeview.pack()
tree_scrollbar.config(command=treeview.yview)

load_data()
#runing a window
root.mainloop()
