import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
import openpyxl as ox
from openpyxl import load_workbook
from tkinter import *
from tkcalendar import Calendar
import os
import locale
import datetime
import numpy as np
from num2words import num2words
import math
from pandastable import Table
import sys

# Invoice
def cust():
    window_cust = tk.Tk()
    window_cust.title("Add Customer")
    window_cust.geometry("1080x780+400+10")
    window_cust.configure(bg='#333333')
    frame_cust=tk.Frame(window_cust,bg='#333333')
    frame_cust.pack(fill=BOTH,expand=True)
    '''
    #vlookup nama cust
    def vlookup(value):
        filter_value = combobox_var.get()
        cust_kode['values'] = [value for value in cust_kode['values'] if isinstance(value, str) and value.startswith(filter_value)]
        result = custfile[custfile[custfile.columns[0]] == value][custfile.columns[1]].values[0]
        return result        
    #vlookup alamat cust
    def addlookup(entry_value):
        filtered_values = [value for value in custfile[custfile.columns[1]].values if isinstance(value, str) and value.startswith(entry_value)]
        if filtered_values:
            result1 = custfile[custfile[custfile.columns[1]] == filtered_values[0]][custfile.columns[2]].values[0]
        else:
            result1 = ""
        return result1
    '''

    #Entry Baru
    def new_entry ():
        cust_kode.delete(0,END)
        nama_entry.delete(0,END)
        alamat_entry.delete(0, END)
        phone_entry.delete(0,END)
        contact_entry.delete(0,END)

    def cust_baru ():
        confirmation1 = messagebox.askyesno("confirmation","apakah data sudah di generate?")
        if confirmation1 :
            cust_kode.delete(0,END)
            nama_entry.delete(0,END)
            alamat_entry.delete(0, END)
            phone_entry.delete(0,END)
            contact_entry.delete(0,END)
        else:
            confirmation2 = messagebox.askyesno("confirmation","apakan anda yakin akan membuat Customer baru?")
            if confirmation2:
                cust_kode.delete(0,END)
                nama_entry.delete(0,END)
                alamat_entry.delete(0, END)
                phone_entry.delete(0,END)
                contact_entry.delete(0,END)
            else :
                return


    file2= ox.load_workbook('Cust_file.xlsx')
    ws2 = file2['Data Base Cust']
    def cust_generate ():  
        kode_cust = cust_kode.get()
        nama_cust = nama_entry.get()
        alamat = alamat_entry.get()
        Phone = phone_entry.get()
        contact_person = contact_entry.get()
        #export ke database invoice
        ws2.append([kode_cust,nama_cust,alamat,Phone,contact_person])
        file2.save('Cust_file.xlsx')
        
        new_entry()
        
    #Design Add Customer
    cust_label = tk.Label(frame_cust,text="Add_Customer",bg='#333333',fg='#FFFFFF',font=["arial",20],justify="center")
    cust_label.pack(side=TOP)
    #Nomor Customer
    no_cust = tk.Label(frame_cust,text="Nomor Id Customer",bg='#333333',fg='#FFFFFF',font=["arial",14],justify="left")
    no_cust.place(x=20,y=120)
    cust_kode = tk.Entry(frame_cust, bg="white",fg="black",font=["arial",12],width=20)
    cust_kode.place(x=220,y=120)

    #Customer
    nama_cust = tk.Label(frame_cust,text="Nama Customer",bg='#333333',fg='#FFFFFF',font=["arial",14],justify="left")
    nama_cust.place(x=20,y=200)
    nama_entry = tk.Entry(frame_cust, bg="white",fg="black",font=["arial",12],width=20)
    nama_entry.place(x=220,y=200)
    alamat_cust = tk.Label(frame_cust,text="Alamat Customer",bg='#333333',fg='#FFFFFF',font=["arial",14],justify="left")
    alamat_cust.place(x=20,y=280)
    alamat_entry = tk.Entry(frame_cust, bg="white",fg="black",font=["arial",12],width=80)
    alamat_entry.place(x=220,y=280)
    phone_cust = tk.Label(frame_cust,text="Phone Customer",bg='#333333',fg='#FFFFFF',font=["arial",14],justify="left")
    phone_cust.place(x=20,y=350)
    phone_entry = tk.Entry(frame_cust, bg="white",fg="black",font=["arial",12],width=20)
    phone_entry.place(x=220,y=350)
    contact_person = tk.Label(frame_cust,text="Contact Person",bg='#333333',fg='#FFFFFF',font=["arial",14],justify="left")
    contact_person.place(x=20,y=420)
    contact_entry = tk.Entry(frame_cust, bg="white",fg="black",font=["arial",12],width=50)
    contact_entry.place(x=220,y=420)

    #generate Button
    generate_cust = tk.Button(frame_cust,text = "Generate Customer", bg="White",fg="black",font=["arial",14],width=40,command=cust_generate)
    generate_cust.place(x=30, y=670)

    #new Customer Button
    new_customer = tk.Button(frame_cust, text = "New Customer", bg="White",fg="black",font=["arial",14],width=40, command=cust_baru)
    new_customer.place(x=30, y=720)


    window_cust.mainloop