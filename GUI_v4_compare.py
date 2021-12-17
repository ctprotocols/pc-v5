# -*- coding: utf-8 -*-
"""
Created on Fri Oct 23 08:18:04 2020

@author: EastmanE
"""
# import os
# import datetime as dt
# import pandas as pd
# import openpyxl as xl
# import numpy as np
# import xml.etree.ElementTree as ET
# from Siemens_settings import columnheaders, finalheaders, renamefields
# import PDFformatter

import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import filedialog
import Siemens_Format_Clean_GUI
import GE_Format_Clean_GUI
import Toshiba_Format_Clean_GUI
import Philips_Format_Clean_GUI
import GE_compare
import Siem_compare
import Tosh_compare
from PIL import ImageTk, Image
import os
import sys

# print('ProtocolCare starting...')
def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
    
def file_browser():
    global filename
    filename = filedialog.askopenfilename(initialdir='/', title = 'select a file', filetypes = (('xlsx files', '*.xlsx'), ('xml files', '*.xml*'), ('all files', '*.')))
    if not filename:
        pass
    else:
        shortname = os.path.split(filename)[1]
        if len(shortname) >12:
            shortname = shortname[-12:]
        label_file_explorer.configure(text = "..."+ shortname + " selected", fg = 'green')

def close():
    window.destroy()

def openhelp():

    helpwin = tk.Toplevel(window)
    helpwin.resizable(width=False, height=False)
    helpwin.geometry("375x375")

    faqimg = Image.open(resource_path('faq.png'))
    faqimg = faqimg.resize((375, 375))

    faq = ImageTk.PhotoImage(faqimg)
    faqlabel = tk.Label(helpwin, image=faq)
    faqlabel.image = faq
    faqlabel.place(x=0, y=0)

    helpwin.mainloop()
    
def go():
    machinename = entry_name.get("1.0",'end-1c')
    typeofscanner = program_name.get()
    label_progress.configure(text = "Processing...")
    label_progress.update()
    if 'GE' in typeofscanner:
        GE_Format_Clean_GUI.GE(filename, machinename, typeofscanner)        
    elif 'Siemens' in typeofscanner:
        Siemens_Format_Clean_GUI.Siemens(filename, machinename, typeofscanner)        
    elif 'Toshiba/Canon' in typeofscanner:
        Toshiba_Format_Clean_GUI.Toshiba(filename, machinename, typeofscanner)        
    elif 'Philips' in typeofscanner:
        Philips_Format_Clean_GUI.Philips(filename, machinename, typeofscanner)
    label_progress.configure(text = "Complete.")
    label_progress.update()


def reformat():
    startwin.destroy()
    translogo = Image.open(resource_path('logotransparent.png'))
    translogo = translogo.resize((375, 600))
    
    
    global window
    window = tk.Tk()
    window.resizable(width=False, height=False)
    window.geometry("375x600")
    # window.configure(bg = "white")
    
    
    #Add logo
    test = ImageTk.PhotoImage(translogo)
    label1 = tk.Label(image=test)
    label1.image = test
    label1.place(x=-2, y=-2)
    
    
    machinename = tk.StringVar()
    machinename.set("None")
    global entry_name
    entry_name = tk.Text(height = 1, width = 25, selectborderwidth = 1, relief = tk.SOLID, font = ("Helvetica", 10), bd = 1) 
    entry_name.place(x=97, y = 207)
    
    
    canonphoto = Image.open(resource_path('go.png'))
    canonphoto = canonphoto.resize((100, 100))
    
    canonphototk = ImageTk.PhotoImage(canonphoto)
    canonbutton = tk.Button(image = canonphototk, relief = tk.FLAT, bg = "#ec1c24", activebackground = "#ec1c24", borderwidth = 0, command = go)
    canonbutton.image = canonphototk
    canonbutton.place(x=135, y=450)
    
    exitphoto = Image.open(resource_path('exit.png'))
    exitphoto = exitphoto.resize((50, 50))
    
    exitphototk = ImageTk.PhotoImage(exitphoto)
    exitbutton = tk.Button(image = exitphototk, relief = tk.FLAT, bg = "#ec1c24", activebackground = "#ec1c24", borderwidth = 0, command = close)
    exitbutton.image = exitphototk
    exitbutton.place(x=50, y=500)
    
    helpphoto = Image.open(resource_path('help.png'))
    helpphoto = helpphoto.resize((50, 50))
    
    helpphototk = ImageTk.PhotoImage(helpphoto)
    helpbutton = tk.Button(image = helpphototk, relief = tk.FLAT, bg = "#ec1c24", activebackground = "#ec1c24", borderwidth = 0, command = openhelp)
    helpbutton.image = helpphototk
    helpbutton.place(x=270, y=500)
    
    homephoto = Image.open(resource_path('home.png'))
    homephoto = homephoto.resize((50, 50))
    
    homephototk = ImageTk.PhotoImage(homephoto)
    homebutton = tk.Button(image = homephototk, relief = tk.FLAT, bg = "#ec1c24", activebackground = "#ec1c24", borderwidth = 0, command = returnhome_r)
    homebutton.image = homephototk
    homebutton.place(x=280, y=20)


    
    filephoto = Image.open(resource_path('fileselect.png'))
    filephoto = filephoto.resize((136, 32))
    filephototk = ImageTk.PhotoImage(filephoto)
    button_explore = tk.Button(command = file_browser, image = filephototk, bg = "white", activebackground = "white", relief = tk.FLAT, borderwidth = 0)  
    button_explore.place(x = 117, y = 340)
    
    global label_file_explorer
    label_file_explorer = tk.Label(text = "No protocol file selected",  fg = "#ec1c24", bg = "white", font = "helvetica") 
    label_file_explorer.place(x=99, y=385)
    
    global label_progress
    label_progress = tk.Label(text = "Press GO!",  fg = "white", bg = "#ec1c24", font = "helvetica") 
    label_progress.place(x=145, y=560)
    
    typelist = ['GE Lightspeed VCT', 'GE Revolution', 'GE Discovery HD 750', 
                'GE Discovery 690', 'GE Optima 560',  'Philips IQon', 'Siemens Biograph Vision', 
                'Siemens Somatom Definition Flash','Siemens Intevo BOLD', 'Siemens Somatom Force', 
                'Toshiba/Canon Aquilion ONE Genesis','Toshiba/Canon Aquilion 64']
    
      
    scannertype = tk.StringVar()
    scannertype.set(typelist[0])
    global program_name
    program_name = Combobox(window, values = typelist, state = 'readonly')
    program_name.current(0)
    program_name.place(x=115, y = 275)
    
    
    window.title('ProtocolCare') 
    title_font_tuple = ("Helvetica", 14, "bold")
    instruct_font_tuple = ("Helvetica", 12, "italic")
    font_tuple = ("Helvetica", 12)
    
    window.mainloop() 
    

def compare():
        
    def old_file_browser():
        global oldfilename
        oldfilename = filedialog.askopenfilename(initialdir='/', title = 'select a file', filetypes = (('xlsx files', '*.xlsx'), ('all files', '*.')))
        if not oldfilename:
            pass
        else:
            shortname = os.path.split(oldfilename)[1]
            if len(shortname) >12:
                shortname = shortname[-12:]
            file1_explorer.configure(text = "..."+ shortname + " selected", fg = 'green')
    
    def new_file_browser():
        global newfilename
        newfilename = filedialog.askopenfilename(initialdir='/', title = 'select a file', filetypes = (('xlsx files', '*.xlsx'), ('all files', '*.')))
        if not newfilename:
            pass
        else:
            shortname = os.path.split(newfilename)[1]
            if len(shortname) >12:
                shortname = shortname[-12:]
            file2_explorer.configure(text = "..."+ shortname + " selected", fg = 'green')

    def compareclose():
        comparewindow.destroy()

    def comparehelp():
    
        helpwin = tk.Toplevel(comparewindow)
        helpwin.resizable(width=False, height=False)

        helpwin.geometry("375x375")
    
        faqimg = Image.open(resource_path('comparehelp.png'))
        faqimg = faqimg.resize((375, 375))
    
        faq = ImageTk.PhotoImage(faqimg)
        faqlabel = tk.Label(helpwin, image=faq)
        faqlabel.image = faq
        faqlabel.place(x=0, y=0)
    
        helpwin.mainloop()
    def compare_go():
        label_progress_c.configure(text = "Processing...")
        label_progress_c.update()

        vendorname = vendor_name.get()
        if vendorname == 'GE':
            GE_compare.GE(oldfilename, newfilename)
        if vendorname == 'Siemens':
            Siem_compare.Siemens(oldfilename, newfilename)
        if vendorname == 'Toshiba/Canon':
            Tosh_compare.Toshiba(oldfilename, newfilename)
        label_progress_c.configure(text = "Complete.")
        label_progress_c.update()

    startwin.destroy()
    translogo = Image.open(resource_path('comparebkg.png'))
    translogo = translogo.resize((375, 600))
    
    global comparewindow
    comparewindow = tk.Tk()
    comparewindow.geometry("375x600")
    comparewindow.resizable(width=False, height=False)
    # window.configure(bg = "white")
    
    
    #Add logo
    test = ImageTk.PhotoImage(translogo)
    label1 = tk.Label(image=test)
    label1.image = test
    label1.place(x=-2, y=-2)
    
    typelist = ['GE', 'Siemens', 'Toshiba/Canon']
    
      
    scannertype = tk.StringVar()
    scannertype.set(typelist[0])
    
    global vendor_name
    vendor_name = Combobox(comparewindow, justify = 'center', values = typelist, state = 'readonly')
    vendor_name.current(0)
    vendor_name.place(x=115, y = 395)

    
    
    canonphoto = Image.open(resource_path('goblue.png'))
    canonphoto = canonphoto.resize((100, 100))
    
    canonphototk = ImageTk.PhotoImage(canonphoto)
    canonbutton = tk.Button(image = canonphototk, relief = tk.FLAT, bg = "#004AAD", activebackground = "#004AAD", borderwidth = 0, command = compare_go)
    canonbutton.image = canonphototk
    canonbutton.place(x=135, y=450)
    
    exitphoto = Image.open(resource_path('exitblue.png'))
    exitphoto = exitphoto.resize((50, 50))
    
    exitphototk = ImageTk.PhotoImage(exitphoto)
    exitbutton = tk.Button(image = exitphototk, relief = tk.FLAT, bg = "#004AAD", activebackground = "#004AAD", borderwidth = 0, command = compareclose)
    exitbutton.image = exitphototk
    exitbutton.place(x=50, y=500)
    
    helpphoto = Image.open(resource_path('helpblue.png'))
    helpphoto = helpphoto.resize((50, 50))
    
    helpphototk = ImageTk.PhotoImage(helpphoto)
    helpbutton = tk.Button(image = helpphototk, relief = tk.FLAT, bg = "#004AAD", activebackground = "#004AAD", borderwidth = 0, command = comparehelp)
    helpbutton.image = helpphototk
    helpbutton.place(x=270, y=500)
    
    homephoto = Image.open(resource_path('home.png'))
    homephoto = homephoto.resize((50, 50))
    
    homephototk = ImageTk.PhotoImage(homephoto)
    homebutton = tk.Button(image = homephototk, relief = tk.FLAT, bg = "#004AAD", activebackground = "#004AAD", borderwidth = 0, command = returnhome_c)
    homebutton.image = homephototk
    homebutton.place(x=280, y=20)

    
    filephoto = Image.open(resource_path('fileselect.png'))
    filephoto = filephoto.resize((136, 32))
    filephototk = ImageTk.PhotoImage(filephoto)
    
    global file1_explorer
    file1_explorer = tk.Label(text = "No protocol file selected",  fg = "#004AAD", bg = "white", font = "helvetica") 
    file1_explorer.place(x=99, y=225)
    global file2_explorer
    file2_explorer = tk.Label(text = "No protocol file selected",  fg = "#004AAD", bg = "white", font = "helvetica") 
    file2_explorer.place(x=99, y=325)
    
    
    old_file_button = tk.Button(image = filephototk, bg = "white", activebackground = "white", relief = tk.FLAT, borderwidth = 0, command=old_file_browser)  
    old_file_button.place(x = 117, y = 190)
    
    
    new_file_button = tk.Button(image = filephototk, bg = "white", activebackground = "white", relief = tk.FLAT, borderwidth = 0, command=new_file_browser)  
    new_file_button.place(x = 117, y = 293)
    
    
    
    global label_progress_c
    label_progress_c = tk.Label(text = "Press GO!",  fg = "white", bg = "#004AAD", font = "helvetica") 
    label_progress_c.place(x=145, y=560)
    
    
      
    
    
    comparewindow.title('ProtocolCare') 
    title_font_tuple = ("Helvetica", 14, "bold")
    instruct_font_tuple = ("Helvetica", 12, "italic")
    font_tuple = ("Helvetica", 12)
    
    comparewindow.mainloop() 

# Let the window wait for any events 
#Startup window
def returnhome_r():
    window.destroy()
    home()
def returnhome_c():
    comparewindow.destroy()
    home()
def home():
    global startwin
    startwin = tk.Tk()
    startwin.geometry("375x600")
    startwin.resizable(width=False, height=False)
    # window.configure(bg = "white")
    
    
    #Add logo
    startbkg = Image.open(resource_path('menu.png'))
    startbkg = startbkg.resize((375, 600))
    
    tkstartbkg = ImageTk.PhotoImage(startbkg)
    startlabel = tk.Label(image=tkstartbkg)
    startlabel.image = tkstartbkg
    startlabel.place(x=-2, y=-2)
    
    #Add logo
    
    reformatbutton = Image.open(resource_path('reformatbutton.png'))
    reformatbutton = reformatbutton.resize((217, 58))
    
    tkreformatbutton = ImageTk.PhotoImage(reformatbutton)
    reformatlabel = tk.Button(image=tkreformatbutton, relief = tk.FLAT, highlightthickness=0, activebackground = "#ec1c24", borderwidth = 0, command = reformat)
    reformatlabel.image = tkreformatbutton
    reformatlabel.place(x=90, y=40)
    
    comparebutton = Image.open(resource_path('comparebutton.png'))
    comparebutton = comparebutton.resize((197, 51))
    
    tkcomparebutton = ImageTk.PhotoImage(comparebutton)
    comparelabel = tk.Button(image=tkcomparebutton, relief = tk.FLAT, highlightthickness=0, activebackground = "#004AAD", borderwidth = 0, command = compare)
    comparelabel.image = tkcomparebutton
    comparelabel.place(x=100, y=510)
    
    # compare()
    
    startwin.mainloop() 
    
home()