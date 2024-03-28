import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import *
import customtkinter as ctk
from Jurnal import jurnal
from Invoice import invoice
from Receipt import receipt
from Trial_Balance import tb
from Neraca import balance_sheet
from Laba_Rugi import laba_rugi
from PIL import ImageTk,Image

#Menu
def create_menu():
    window_menu = tk.Tk()
    window_menu.title("Menu")
    window_menu.geometry('400x700+10+20')
    window_menu.configure(bg='#333333')    
    frame_menu=tk.Frame(bg='#333333')

#Design Form Menu
#Form menu
    
    input_jurnal = tk.Button(frame_menu, text="Input Jurnal",bg="white", fg="black",font=["arial",18],command=(jurnal))
    input_invoice = tk.Button(frame_menu, text="Invoice",bg="white", fg="black",font=["arial",18],command=(invoice))
    input_receipt = tk.Button(frame_menu, text="Receipt",bg="white", fg="black",font=["arial",18],command=(receipt))
    trial_balance = tk.Button(frame_menu, text="Neraca Saldo",bg="white", fg="black",font=["arial",18],command=(tb))
    neraca_button = tk.Button(frame_menu, text="Neraca",bg="white", fg="black",font=["arial",18],command=(balance_sheet))
    lr_button = tk.Button(frame_menu, text="laba_rugi",bg="white", fg="black",font=["arial",18], command=(laba_rugi))
    cash_flow = tk.Button(frame_menu, text="Arus Kas",bg="white", fg="black",font=["arial",18])
    input_rasio = tk.Button(frame_menu, text="Rasio",bg="white", fg="black",font=["arial",18])
    aging_piutang = tk.Button(frame_menu, text="Umur Piutang",bg="white", fg="black",font=["arial",18])
    add_cust = tk.Button(frame_menu, text="Add Customer",bg="white", fg="black",font=["arial",18])
    add_akun = tk.Button(frame_menu, text="Add Akun",bg="white", fg="black",font=["arial",18])

#Penempatan
    input_jurnal.grid(row=3,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    input_invoice.grid(row=4,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    input_receipt.grid(row=5,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    trial_balance.grid(row=6,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    lr_button.grid(row=7,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    neraca_button.grid(row=8,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    cash_flow.grid(row=9,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    input_rasio.grid(row=10,column=1, columnspan=2, pady=20,sticky="W",padx=10)
    add_cust.grid(row=4,column=6, columnspan=2, pady=20,sticky="W",padx=10)
    add_akun.grid(row=3,column=6, columnspan=2, pady=20,sticky="W",padx=10)
    aging_piutang.grid(row=5,column=6, columnspan=2, pady=20,sticky="W",padx=10)
    
    frame_menu.pack()
    window.mainloop()