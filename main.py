import openpyxl as op
import tkinter as tk
from tkinter import ttk

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
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insira os dados")
widgets_frame.grid(column=0, row=0, padx=700, pady=300)

#input informations
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Produto...")
name_entry.bind("<FocusIn>", lambda event:name_entry.delete(0, "end"))
name_entry.grid(column=1, row=0, padx=5, pady=5, sticky="ew")

#runing a window
root.mainloop()



