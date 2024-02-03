import openpyxl
import tkinter as tk
from tkinter import ttk

#backend
def load_data():
    path = r"C:\Users\gabri\Documents\Almoxarifado\Almoxarifado\bd.xlsx"
    wb = openpyxl.load_workbook(path)
    sheet = wb.active

    list_values = list(sheet.values)
    print(list_values)

    for col_name in cols:
        treeview.heading(col_name, text=)
def insert_data():
    nome = name_entry.get()
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

def mode_theme():
    if mode_switch.instate(["selected"]):
        style.theme_use('forest-light')
    else:
        style.theme_use('forest-dark')

#frontend
#stating a window
root = tk.Tk()
root.title("Almoxarifado")

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
name_entry = ttk.Entry(widgets_frame, width=40)
name_entry.insert(0, "Produto...")
name_entry.bind("<FocusIn>", lambda event:name_entry.delete(0, "end"))
name_entry.grid(column=0, row=0, padx=5, pady=5, sticky="ew", columnspan=2)

num_spinbox = ttk.Spinbox(widgets_frame, from_=0, to =10000, width=6)
num_spinbox.insert(0, "Qnt")
num_spinbox.grid(column=3, row=0, padx=5, pady=5, sticky="ew")

dep = ["Cerpat","Cerest", "VIEP", "VISA", "Ambiental", "Laboratório", "Zoonoses", "Endemias"]
dep_pack = ttk.Combobox(widgets_frame, values=dep, width=30)
dep_pack.set("Departamentos...")
dep_pack.grid(column=0, row=1, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="Teste Covid", variable=a)
checkbutton.grid(column=4, row=0, padx=5, pady=5,sticky="nsew")

button = ttk.Button(widgets_frame, text="Inserir Dados".upper(), command=insert_data)
button.grid(column=0, row=2, padx=5, pady=5,sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid (column=0, row=3, padx=5, pady=10, sticky="nsew",columnspan=5)

mode_switch = ttk.Checkbutton(widgets_frame, text="Modo", style="Switch", command=mode_theme)
mode_switch.grid(column=0, row=6, padx=5, pady=10, sticky="nsew")

#Creating previwer
tree_frame = ttk.Frame(frame)
tree_frame.grid(column=1, row=0, padx=(0,20), pady=(10))
tree_scrollbar = ttk.Scrollbar(tree_frame)
tree_scrollbar.pack(side="right", fill="y")

cols = ("Produto","Departamento","Quantidade","Valor unitário")
treeview = ttk.Treeview(tree_frame, show="headings", yscrollcommand=tree_scrollbar.set, columns = cols, height=13)

treeview.column("Produto", width=100)
treeview.column("Departamento", width=100)
treeview.column("Quantidade", width=50)
treeview.column("Valor unitário", width=50)
treeview.pack()
tree_scrollbar.config(command=treeview.yview)

#runing a window
root.mainloop()



