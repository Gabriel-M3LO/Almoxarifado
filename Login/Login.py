import customtkinter
import tkinter as tk
from tkinter import ttk

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

janela = customtkinter.CTk()
janela.geometry("720x480")
janela.title("Login")
janela.resizable(width=False, height=False)

#frame
frame = customtkinter.CTkFrame(master=janela, width=287, height=480, fg_color="#0091B1",curve=0)
frame.pack(side="left")
#frame.config(background="blue")

janela.mainloop()