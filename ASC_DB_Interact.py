#!/usr/bin/python3

# ASC Database Interact
# Author: B. Young
# Last Modified: May 27, 2025
# Contact Info:
# Email: brent.young@wilkes.edu
# Office: SLC 406
# Ext: 3824

#packages to import
import tkinter as tk
from tkinter import scrolledtext
from tkinter import filedialog as fd
from tkinter import ttk
from collections import namedtuple
from functools import partial
import mysql.connector 
import sys
import os
from array import *
from datetime import datetime
import csv

#path to icon file (UPDATE THIS)
icon_path = '/path/to/Wilkes_W.png'

#list of accept keys:
acc_keys = ['<Return>', '<KP_Enter>']

#standard header for LATEX files
stdhd = ["\\documentclass[12pt]{article}\n","\\pagestyle{empty}\n","\\usepackage{fontspec}\n", "\\setmainfont{DejaVu Serif}\n", "\\usepackage{extsizes}\n","\\usepackage{amsfonts}\n","\\usepackage{amsmath}\n","\\usepackage{amsthm}\n","\\usepackage{amssymb}\n",
       "\\usepackage{mathrsfs}\n","\\usepackage{comment}\n","\\usepackage{adjustbox}\n","\\usepackage{hyperref}\n","\\addtolength{\\hoffset}{-65pt}\n",
       "\\addtolength{\\voffset}{-65pt}\n", "\\addtolength{\\textwidth}{125pt}\n","\\addtolength{\\textheight}{100pt}\n","\\addtolength{\\marginparwidth}{-100pt}\n"]

#function to escape any special Latex characters in a string
def texify(string):
	string = string.replace("{", r"\{")
	string = string.replace("}", r"\}")
	string = string.replace("\\",r'\textbackslash{}' )
	string = string.replace("\n", r"\newline ")
	string = string.replace("&", r"\&")
	string = string.replace("_", r"\_")
	string = string.replace("^", r"\^{}")
	string = string.replace("$", r"\$")
	string = string.replace("%", r"\%")
	string = string.replace("#", r"\#")
	string = string.replace("~", r"\~{}")
	string = string.replace('"',r'\verb|"|')
	string = string.replace("'",r"\verb|'|")
	return string

#control button size
buttonsz=20

#login routine
logstate = 0
pswd = " "

#login connection function
def dblogin(event):
    global logstate
    global cnx
    global pswd
    pswd = psswd.get()
    loginfo = {
        'user': 'root',
        'password': pswd,
        'host': 'localhost',
        'database': 'Appeals_DB',
        'auth_plugin':'mysql_native_password',
        'raise_on_warnings': True}
    try:
        cnx = mysql.connector.connect(**loginfo)
    except mysql.connector.Error as err:
        if err.errno == 1045:
            errorDialog = tk.Toplevel(login)
            errorDialog.wm_title("Log In Error")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path )
            errorDialog.iconphoto(False, ph1)
            tk.Label(errorDialog,text="Incorrect Database Password!",font="Helvetica 12 bold").grid(row = 0, column = 0)
            okButton = tk.Button(errorDialog, text = "OK", command=errorDialog.destroy)
            okButton.grid(row = 3, column = 1, padx = 5, pady = 5)
            for key in acc_keys:
                errorDialog.bind(key, lambda event: errorDialog.destroy())
        elif err.errno == 1049:
            os.popen("mysql -u root -p"+pswd+" -e \"CREATE DATABASE Appeals_DB\"")
            errorDialog = tk.Toplevel(login)
            errorDialog.wm_title("Log In Error")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path )
            errorDialog.iconphoto(False, ph1)
            tk.Label(errorDialog,text="Appeals_DB Database empty!",font="Helvetica 12 bold").grid(row = 0, column = 0)
            tk.Label(errorDialog,text="You must restore from a previous backup!",font="Helvetica 12 bold").grid(row = 2, column = 0)
            okButton = tk.Button(errorDialog, text = "OK", command=errorDialog.destroy)
            okButton.grid(row = 5, column = 1, padx = 5, pady = 5)
            for key in acc_keys:
                errorDialog.bind(key, lambda event: errorDialog.destroy())
        else:
            errorDialog = tk.Toplevel(login)
            errorDialog.wm_title("Log In Error")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path )
            errorDialog.iconphoto(False, ph1)
            tk.Label(errorDialog,text="Error Number: " + str(err.errno),font="Helvetica 12 bold").grid(row = 0, column = 0)
            tk.Label(errorDialog,text=err,font="Helvetica 10").grid(row = 1, column = 0)
            okButton = tk.Button(errorDialog, text = "OK", command=errorDialog.destroy)
            okButton.grid(row = 5, column = 1, padx = 5, pady = 5)
            for key in acc_keys:
                errorDialog.bind(key, lambda event: errorDialog.destroy())
            
    else:
        logstate = 1
        login.destroy()
        
def closer(event):
    quit()

#login popup window
while logstate == 0:       
    login = tk.Tk(className="asc db")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path )
    login.iconphoto(False, ph1)
    login.geometry("450x100")
    login.wm_title("ASC Database Login")
    tk.Label(login,text="Enter DB Password:",font="Helvetica 12 bold").grid(row = 0, column = 0)
    psswd = tk.Entry(login, show='*', width=25,font=("Helvetica", 15))
    psswd.grid(row = 0, column = 3, padx = 5, pady = 5)
    okButton = tk.Button(login, text = "Login")
    okButton.bind('<Button-1>', dblogin)
    okButton.grid(row = 3, column = 0, padx = 5, pady = 5)
    cancelButton = tk.Button(login, text = "Cancel")
    cancelButton.bind('<Button-1>', closer)
    cancelButton.grid(row = 3, column = 3, padx = 5, pady = 5)
    for key in acc_keys:
        login.bind(key, dblogin)
    psswd.focus_set()
    login.mainloop()

#make sure dates (in YYYY-MM-DD format) are sensible
def date_clean(date):
    max_yr = 2999
    min_yr = 2000
    x = date.split("-")
    if len(x) != 3:
        x = ['2000', '01', '01']
    x[0]=int(x[0])
    x[1]=int(x[1])
    x[2]=int(x[2])
    month_days={1:31, 2:29, 3:31, 4:30, 5:31, 6:30, 7:31, 8:31, 9:30, 10:31, 11:30, 12:31}
    if x[0] < min_yr:
        x[0] = min_yr
    if x[0] > max_yr:
        x[0] = min_yr
    if x[1] > 12:
        x[1] = 12
    if month_days[x[1]] < x[2]:
        x[2] = month_days[x[1]]
    x[0] = str(x[0])
    if x[1] > 9:
        x[1] = str(x[1])
    else:
        x[1] = '0'+str(x[1])
    if x[2] > 9:
        x[2] = str(x[2])
    else:
        x[2] = '0'+str(x[2])
    return x[0]+"-"+x[1]+"-"+x[2]

#take date string in format YYYY-MM-DD and print a fancy date string
def fancy_date(raw_date):
    months=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    x = raw_date.split("-")
    ret = months[int(x[1])-1] + " " + x[2] + ", " + x[0]
    return ret
    
#print WIN in easier format to read
#assumes WIN is an integer like 900123456
def fancy_WIN(win):
    win = str(win)
    return win[0:3] + "-" + win[3:5] + "-" + win[5:9]
    

#compare two dates in YYYY-MM-DD format. Returns 0 if dates are sequential, 1 otherwise
def date_comp(date1, date2):
    x = date1.split("-")
    y = date2.split("-")
    if int(x[0]) > int(y[0]):
        return 1
    if (int(x[0]) == int(y[0]))&(int(x[1]) > int(y[1])):
        return 1
    if (int(x[0]) == int(y[0]))&(int(x[1]) == int(y[1]))&(int(x[2]) > int(y[2])):
        return 1
    return 0

#dbEntry class to hold DB entries
class dbEntry(object):
    __slots__ = ('win','date', 'time', 'first_name','last_name','comm', 'nt','mins','motion','dec','readmit')
    def __init__(self):
        self.win = 900000000
        self.date = '1900-01-01'
        self.time = '13:59'
        self.first_name = ''
        self.last_name = ''
        self.comm = ''
        self.nt = ''
        self.mins = ''
        self.motion = ''
        self.dec = ''
        self.readmit = ''
        
#namedtuple for verification of entries
Verifier = namedtuple('Verifier',['win','first','last','date','mins','readmit'])

#DB entry verification function
def verifyFields(dbentry):
    try:
        if (int(dbentry.win) > 900000000) & (int(dbentry.win) < 900999999):
            wvfy = 0
        else:
            wvfy = 1
    except:
        wvfy = 1

    if (dbentry.first_name == '') or (dbentry.first_name == ' ') or (dbentry.first_name == 'First Name?'):
        fvfy = 1
    else:
        fvfy = 0

    if (dbentry.last_name == '') or (dbentry.last_name == ' ') or (dbentry.last_name == 'Last Name?'):
        lvfy = 1
    else:
        lvfy = 0

    if dbentry.date == '2000-01-01':
        dvfy = 1
    else:
        dvfy = 0

    if (dbentry.mins == '\n') or (dbentry.mins == 'Insert minutes?\n') or (dbentry.mins == ''):
        mvfy = 1
    else:
        mvfy = 0

    if (dbentry.readmit == 'Y') or (dbentry.readmit == 'N'):
        rvfy = 0
    else:
        rvfy = 1
    return Verifier(wvfy, fvfy, lvfy, dvfy, mvfy, rvfy)

#DB entry verification form
def verifyEntry(dbentry):
    vfy = verifyFields(dbentry)
    def colorizer(flg):
        if flg == 0: return 'black'
        else: return 'red'

    def errortext(vfy, fld, dbentry):
        if (fld == 'win') & (vfy.win == 1):
            return '900000000'
        elif (fld == 'win') & (vfy.win == 0):
            return dbentry.win

        if (fld == 'first') & (vfy.first == 1):
            return 'First Name?'
        elif (fld == 'first') & (vfy.first == 0):
            return dbentry.first_name

        if (fld == 'last') & (vfy.last == 1):
            return 'Last Name?'
        elif (fld == 'last') & (vfy.last == 0):
            return dbentry.last_name
        
        if (fld == 'mins') & (vfy.mins == 1):
            return 'Insert minutes?'
        elif (fld == 'mins') & (vfy.mins == 0):
            return dbentry.mins

        if (fld == 'readmit') & (vfy.readmit == 1):
            return '?'
        elif (fld == 'readmit') & (vfy.readmit == 0):
            return dbentry.readmit

    def submit():
        newEntry = dbEntry()
        newEntry.win = WIN.get()
        newEntry.date = date_clean(DATE.get())
        newEntry.time = time_entry.get()
        newEntry.first_name = FN.get()
        newEntry.last_name = LN.get()
        newEntry.comm = comm_entry.get('1.0', tk.END)
        newEntry.nt = nt_entry.get()
        newEntry.mins = min_entry.get('1.0', tk.END)
        newEntry.motion = mot_entry.get('1.0', tk.END)
        newEntry.dec = dec_entry.get('1.0', tk.END)
        newEntry.readmit = RA.get()
        entryWindow.destroy()
        #if required fields all good:  add to DB.  otherwise, restart verification process with bad fields in red
        if verifyFields(newEntry)==(0,0,0,0,0,0):
            cmd = "REPLACE INTO Appeal_Records (WIN, First_Name, Last_Name, Date, Time, Committee_Members, Note_Taker, Minutes, Motion, Decision, Readmit) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
            try:
                curs = cnx.cursor()
                curs.execute(cmd,(str(newEntry.win), newEntry.first_name, newEntry.last_name, newEntry.date, newEntry.time, newEntry.comm, newEntry.nt, newEntry.mins, newEntry.motion, newEntry.dec, newEntry.readmit))
                cnx.commit()
                successDialog = tk.Toplevel(root)
                successDialog.wm_title("DB Entry Success")
                #application icon
                ph1 = tk.PhotoImage(file = icon_path )
                successDialog.iconphoto(False, ph1)
                tk.Label(successDialog,text="DB Entry for WIN = " + str(newEntry.win) + " was successful!",font="Helvetica 12 bold").grid(row = 0, column = 0)
                closeButton = tk.Button(successDialog, text = "Close", command=successDialog.destroy)
                closeButton.grid(row = 3, column = 0, padx = 5, pady = 5)
                againButton = tk.Button(successDialog, text = "Add Another", command=lambda:[successDialog.destroy(), createEntry()])
                againButton.grid(row = 3, column = 3, padx = 5, pady = 5)
                
            except mysql.connector.Error as err:
                errorDialog = tk.Toplevel(root)
                errorDialog.wm_title("DB Entry Error")
                #application icon
                ph1 = tk.PhotoImage(file = icon_path)
                errorDialog.iconphoto(False, ph1)
                tk.Label(errorDialog,text="Error Number: " + str(err.errno),font="Helvetica 12 bold").grid(row = 0, column = 0)
                tk.Label(errorDialog,text=err,font="Helvetica 10").grid(row = 1, column = 0)
                okButton = tk.Button(errorDialog, text = "OK", command=errorDialog.destroy)
                okButton.grid(row = 5, column = 1, padx = 5, pady = 5)
        else:
            verifyEntry(newEntry)

    #create verification window
    entryWindow = tk.Toplevel(root)
    entryWindow.grab_set()
    entryWindow.geometry("800x800")
    entryWindow.wm_title("Entry Verification Window")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    entryWindow.iconphoto(False, ph1)
    tk.Label(entryWindow,text="WIN (format 900123456):").grid(row = 0, column = 0,sticky=tk.E)
    WIN=tk.Entry(entryWindow, width=15, font=("Helvetica", 15), fg = colorizer(vfy.win))
    WIN.insert(0,errortext(vfy, 'win', dbentry))
    WIN.grid(row = 0, column = 3, columnspan = 2, sticky = tk.W+tk.E, padx = 2, pady = 2)
    WIN.focus()
    tk.Label(entryWindow,text="Date (format YYYY-MM-DD):").grid(row = 0, column = 6,sticky=tk.E)
    DATE=tk.Entry(entryWindow, width=15, font=("Helvetica", 15),fg = colorizer(vfy.date))
    DATE.grid(row = 0, column = 9, padx = 2, pady = 2)
    DATE.insert(0,dbentry.date)
    tk.Label(entryWindow,text="Time (HH:MM in 24 hour format):").grid(row = 3, column = 6,sticky=tk.E)
    time_entry=tk.Entry(entryWindow, width=8, font=("Helvetica", 15))
    time_entry.grid(row = 3, column = 9, padx = 2, pady = 2,sticky=tk.W)
    time_entry.insert(0,dbentry.time)
    tk.Label(entryWindow,text="First Name:").grid(row = 6, column = 0,sticky=tk.E)
    FN=tk.Entry(entryWindow, width=15, font=("Helvetica", 15), fg = colorizer(vfy.first))
    FN.grid(row = 6, column = 3, columnspan = 3, sticky = tk.W+tk.E, padx = 2, pady = 2)
    FN.insert(0,errortext(vfy, 'first', dbentry))
    tk.Label(entryWindow,text="Last Name:").grid(row = 6, column = 6,sticky=tk.E)
    LN=tk.Entry(entryWindow, width=15, font=("Helvetica", 15), fg = colorizer(vfy.last))
    LN.grid(row = 6, column = 9, padx = 2, pady = 2)
    LN.insert(0,errortext(vfy, 'last', dbentry))
    tk.Label(entryWindow,text="Committee Members:").grid(row = 9, column = 0,sticky=tk.E)
    comm_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=3,font=("Helvetica", 15))
    comm_entry.grid(column=3, row=9, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    comm_entry.insert('1.0',dbentry.comm)
    tk.Label(entryWindow,text="Note Taker:").grid(row = 12, column = 0,sticky=tk.E)
    nt_entry=tk.Entry(entryWindow, width=3, font=("Helvetica", 15))
    nt_entry.grid(row = 12, column = 3, padx = 2, pady = 2, sticky = tk.W)
    nt_entry.insert(0,dbentry.nt)
    tk.Label(entryWindow,text="Minutes:").grid(row = 15, column = 0,sticky=tk.E)
    min_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=15,font=("Helvetica", 15),fg = colorizer(vfy.mins))
    min_entry.grid(column=3, row=15, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    min_entry.insert('1.0',errortext(vfy, 'mins', dbentry))
    tk.Label(entryWindow,text="Motion:").grid(row = 18, column = 0,sticky=tk.E)
    mot_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=3,font=("Helvetica", 15))
    mot_entry.grid(column=3, row=18, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    mot_entry.insert('1.0',dbentry.motion)
    tk.Label(entryWindow,text="Decision:").grid(row = 21, column = 0,sticky=tk.E)
    dec_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=3,font=("Helvetica", 15))
    dec_entry.grid(column=3, row=21, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    dec_entry.insert('1.0',dbentry.dec)
    tk.Label(entryWindow,text="Readmit?:").grid(row = 24, column = 0,sticky=tk.E)
    RA=tk.Entry(entryWindow, width=2, font=("Helvetica", 15),fg = colorizer(vfy.readmit))
    RA.grid(row = 24, column = 3, padx = 2, pady = 2, sticky = tk.W)
    RA.insert(0,errortext(vfy, 'readmit', dbentry))
    subButton = tk.Button(entryWindow, text = "Submit", command = submit, fg = "red")
    subButton.grid(row = 24, column = 6)
    cancelButton = tk.Button(entryWindow, text = "Cancel", command = entryWindow.destroy)
    cancelButton.grid(row = 25, column = 6)


#DB form entry function
def createEntry():


    #DB submission function
    def submit():
        newEntry = dbEntry()
        newEntry.win = WIN.get()
        newEntry.date = date_clean(DATE.get())
        newEntry.time = time_entry.get()
        newEntry.first_name = FN.get()
        newEntry.last_name = LN.get()
        newEntry.comm = comm_entry.get('1.0', tk.END)
        newEntry.nt = nt_entry.get()
        newEntry.mins = min_entry.get('1.0', tk.END)
        newEntry.motion = mot_entry.get('1.0', tk.END)
        newEntry.dec = dec_entry.get('1.0', tk.END)
        newEntry.readmit = RA.get()
        entryWindow.destroy()
        verifyEntry(newEntry)

    #create entry window
    entryWindow = tk.Toplevel(root)
    entryWindow.grab_set()
    entryWindow.geometry("800x800")
    entryWindow.wm_title("New Entry Window")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    entryWindow.iconphoto(False, ph1)
    tk.Label(entryWindow,text="WIN (format 900123456):").grid(row = 0, column = 0,sticky=tk.E)
    WIN=tk.Entry(entryWindow, width=15, font=("Helvetica", 15))
    WIN.grid(row = 0, column = 3, columnspan = 2, sticky = tk.W+tk.E, padx = 2, pady = 2)
    WIN.focus()
    tk.Label(entryWindow,text="Date (format YYYY-MM-DD):").grid(row = 0, column = 6,sticky=tk.E)
    DATE=tk.Entry(entryWindow, width=15, font=("Helvetica", 15))
    DATE.grid(row = 0, column = 9, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Time (HH:MM in 24 hour format):").grid(row = 3, column = 6,sticky=tk.E)
    time_entry=tk.Entry(entryWindow, width=8, font=("Helvetica", 15))
    time_entry.grid(row = 3, column = 9, padx = 2, pady = 2,sticky=tk.W)
    tk.Label(entryWindow,text="First Name:").grid(row = 6, column = 0,sticky=tk.E)
    FN=tk.Entry(entryWindow, width=15, font=("Helvetica", 15))
    FN.grid(row = 6, column = 3, columnspan = 3, sticky = tk.W+tk.E, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Last Name:").grid(row = 6, column = 6,sticky=tk.E)
    LN=tk.Entry(entryWindow, width=15, font=("Helvetica", 15))
    LN.grid(row = 6, column = 9, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Committee Members:").grid(row = 9, column = 0,sticky=tk.E)
    comm_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=3,font=("Helvetica", 15))
    comm_entry.grid(column=3, row=9, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Note Taker:").grid(row = 12, column = 0,sticky=tk.E)
    nt_entry=tk.Entry(entryWindow, width=3, font=("Helvetica", 15))
    nt_entry.grid(row = 12, column = 3, padx = 2, pady = 2,sticky=tk.W)
    tk.Label(entryWindow,text="Minutes:").grid(row = 15, column = 0,sticky=tk.E)
    min_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=15,font=("Helvetica", 15))
    min_entry.grid(column=3, row=15, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Motion:").grid(row = 18, column = 0,sticky=tk.E)
    mot_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=3,font=("Helvetica", 15))
    mot_entry.grid(column=3, row=18, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Decision:").grid(row = 21, column = 0,sticky=tk.E)
    dec_entry = scrolledtext.ScrolledText(entryWindow, wrap=tk.WORD, width=50, height=3,font=("Helvetica", 15))
    dec_entry.grid(column=3, row=21, columnspan = 20, sticky = tk.W+tk.E, padx = 2, pady = 2)
    tk.Label(entryWindow,text="Readmit?:").grid(row = 24, column = 0,sticky=tk.E)
    RA=tk.Entry(entryWindow, width=2, font=("Helvetica", 15))
    RA.grid(row = 24, column = 3, padx = 2, pady = 2, sticky=tk.W)
    subButton = tk.Button(entryWindow, text = "Proceed", command = submit)
    subButton.grid(row = 24, column = 6)
    cancelButton = tk.Button(entryWindow, text = "Cancel", command = entryWindow.destroy)
    cancelButton.grid(row = 25, column = 6)

#REPORT GENERATION (all functions will return a pdf file compiled using LATEX)

#win report generator function
    
def winReport(data):
    #data is assumed to be a list: first entry is common WIN for group of minutes to follow.
    
    #function to generate report after name of report is entered
    def getFN():
        filename = fnd.get()
        if len(filename) > 1:
            filename1 = filename + ".tex"
            askDialog.destroy()
            directory = fd.askdirectory()
            if directory == () or directory == '':
                return 1
            full_name = os.path.join(directory,filename1)
            #open tex file for writing
            file = open(full_name,"w")
        
            #write coverpage
            for line in stdhd:
                file.write(line)
            file.write("\\title{Academic Standards Committee\\\\{\\huge \\textbf{SPECIFIC WINs REPORT}}}\n")
            file.write("\\begin{document}\n")
            file.write("\\maketitle\n")
            file.write("\\thispagestyle{empty}\n")
            file.write("{\\Large \\textbf{Students Included:}}\n")
            file.write("\\begin{itemize}\n")
            for entry in data:
                file.write("\\item " + fancy_WIN(entry[0]) + ": " + entry[1][0][1] + " " + entry[1][0][2] +" (number of records: " + str(len(entry[1])) + ")" + "\n")
            file.write("\\end{itemize}\n")

            #write individual entries to file
            for win in data:
                for entry in win[1]:
                    file.write("\\newpage \n")
                    file.write("\\noindent \\textbf{WIN: }" + fancy_WIN(entry[0]) + " \n")
                    file.write(" \n")
                    file.write("\\noindent \\textbf{Student Name: } " + texify(entry[1]) + " " + texify(entry[2]) +" \n")
                    file.write(" \n")
                    file.write("\\noindent \\textbf{Date: }" + fancy_date(str(entry[3])) + " \\textbf{Time: }"+ entry[4] +" \n")
                    file.write("\n")
                    file.write("\\smallskip \n")
                    file.write("\n")
                    file.write("\\noindent \\textbf{Committee Members:} \n")
                    file.write(" \n")
                    file.write("\\noindent " + texify(entry[5]) +" \n")
                    file.write(" \n")
                    file.write("\\noindent \\textbf{Note Taker: }" + texify(entry[6]) + " \n")
                    file.write("\n")
                    file.write("\\smallskip \n")
                    file.write("\n")
                    file.write("\\noindent \\textbf{Minutes:} \n")
                    file.write(" \n")
                    file.write("\\noindent " + texify(entry[7]) + " \n")
                    file.write(" \n")
                    file.write("\n")
                    file.write("\\smallskip \n")
                    file.write("\n")
                    file.write("\\noindent \\textbf{Motion: } " + texify(entry[8]) + " \n")
                    file.write(" \n")
                    file.write("\\noindent \\textbf{Decision: } " + texify(entry[9]) + " \n")
                    file.write(" \n")
                    file.write("\\noindent \\textbf{Readmit: }" + entry[10] + "\n")
            file.write("\\end{document}")
            file.close()

            #write shell script to compile tex to pdf and erase auxiliary files
            shscr = os.path.join(directory,"tmp.sh")
            file = open(shscr,"w")
            file.write("#!/bin/sh \n")
            file.write("/usr/bin/xelatex -output-directory " + directory + " " + full_name + "\n")
            file.write("/usr/bin/xelatex -output-directory " + directory + " " + full_name + "\n")
            file.write("rm " + full_name + " \n")
            file.write("rm " + os.path.join(directory,filename + ".log ") + "\n")
            file.write("rm " + os.path.join(directory,filename + ".aux ") + "\n")
            file.write("rm " + os.path.join(directory,filename + ".out ") + "\n")
            #file.write("rm " + os.path.join(directory,"texput.log") + "\n")
            file.write("rm " + shscr)
            file.close()
        
            #run shell script
            os.system("sh " + shscr)

            #success dialog!
            outDialog = tk.Toplevel(root)
            outDialog.wm_title("WIN Report")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            outDialog.iconphoto(False, ph1)
            outDialog.grab_set()
            tk.Label(outDialog,text= "WIN Report: " + os.path.join(directory,filename + ".pdf") + " created!",anchor="w",font="Helvetica 12 bold").grid(row = 0, column = 0)
            okButton = tk.Button(outDialog, text = "OK",command = outDialog.destroy)
            okButton.grid(row = 2, column = 2, padx = 5, pady = 5)
            for key in acc_keys:
                outDialog.bind(key, lambda event: outDialog.destroy())
        
        

    #ask for filename of report and ask for directory to store it in
    askDialog = tk.Toplevel(root)
    askDialog.geometry("550x100")
    askDialog.wm_title("Filename for WIN Report")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    askDialog.iconphoto(False, ph1)
    tk.Label(askDialog,text="Name of report (no extension):",font="Helvetica 12").grid(row = 0, column = 0)
    fnd = tk.Entry(askDialog, width=25,font=("Helvetica", 15))
    fnd.grid(row = 0, column = 3, padx = 5, pady = 5)
    okButton = tk.Button(askDialog, text = "Accept", command=getFN)
    okButton.grid(row = 3, column = 3, padx = 5, pady = 5)
    cancelButton = tk.Button(askDialog, text = "Cancel", command=askDialog.destroy)
    cancelButton.grid(row = 3, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        fnd.bind(key, lambda event: getFN())
    fnd.focus_set()

#date report generator function
def dateReport(startdate, enddate, data):
    #data is assumed to be a list in order by Date:Time
    #the dates are assumed to be strings

    #function to generate report after name of report is entered
    def getFN():
        filename = fnd.get()
        if len(filename) > 1:
            filename1 = filename + ".tex"
            askDialog.destroy()
            directory = fd.askdirectory()
            if directory == () or directory == '':
                return 1
            full_name = os.path.join(directory,filename1)
            #open tex file for writing
            file = open(full_name,"w")
        
            #write coverpage
            for line in stdhd:
                file.write(line)
            file.write("\\title{Academic Standards Committee\\\\{\\huge \\textbf{DATE RANGE REPORT}}}\n")
            file.write("\\begin{document}\n")
            file.write("\\maketitle\n")
            file.write("\\thispagestyle{empty}\n")
            file.write("\\begin{tabular}{rl}\n")
            file.write("\\noindent {\\Large \\textbf{Start Date: } } & {\\Large " + startdate + "} \\\\ \n")
            file.write("\\noindent {\\Large \\textbf{End Date: } } & {\\Large " + enddate + "} \n")
            file.write("\\end{tabular} \n")
        
            #write individual entries to file
            for entry in data:
                file.write("\\newpage \n")
                file.write("\\noindent \\textbf{WIN: }" + fancy_WIN(entry[0]) + " \n")
                file.write(" \n")
                file.write("\\noindent \\textbf{Student Name: }" + texify(entry[1]) + " " + texify(entry[2]) +" \n")
                file.write(" \n")
                file.write("\\noindent \\textbf{Date: }" + fancy_date(str(entry[3])) + " \\textbf{Time: }"+ entry[4] +" \n")
                file.write(" \n")
                file.write("\n")
                file.write("\\smallskip \n")
                file.write("\n")
                file.write("\\noindent \\textbf{Committee Members:}\n")
                file.write(" \n")
                file.write("\\noindent " + texify(entry[5]) +" \n")
                file.write(" \n")
                file.write("\\noindent \\textbf{Note Taker: }" + texify(entry[6]) + " \n")
                file.write(" \n")
                file.write("\n")
                file.write("\\smallskip \n")
                file.write("\n")
                file.write("\\noindent \\textbf{Minutes:} \n")
                file.write(" \n")
                file.write("\\noindent " + texify(entry[7]) + " \n")
                file.write(" \n")
                file.write("\n")
                file.write("\\smallskip \n")
                file.write("\n")
                file.write("\\noindent \\textbf{Motion: }" + texify(entry[8]) + " \n")
                file.write(" \n")
                file.write("\\noindent \\textbf{Decision: }" + texify(entry[9]) + " \n")
                file.write(" \n")
                file.write("\\noindent \\textbf{Readmit: }" + texify(entry[10]) + "\n")
            file.write("\\end{document}")
            file.close()

            #write shell script to compile tex to pdf and erase auxiliary files
            shscr = os.path.join(directory,"tmp.sh")
            file = open(shscr,"w")
            file.write("#!/bin/sh \n")
            file.write("/usr/bin/xelatex -output-directory " + directory + " " + full_name + "\n")
            file.write("/usr/bin/xelatex -output-directory " + directory + " " + full_name + "\n")
            file.write("rm " + full_name + " \n")
            file.write("rm " + os.path.join(directory,filename + ".log ") + "\n")
            file.write("rm " + os.path.join(directory,filename + ".aux ") + "\n")
            file.write("rm " + os.path.join(directory,filename + ".out ") + "\n")
            #file.write("rm " + os.path.join(directory,"texput.log") + "\n")
            file.write("rm " + shscr)
            file.close()
        
            #run shell script
            os.system("sh " + shscr)

            #success dialog!
            outDialog = tk.Toplevel(root)
            outDialog.wm_title("Date Range Report")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            outDialog.iconphoto(False, ph1)
            outDialog.grab_set()
            tk.Label(outDialog,text= "Date Range Report: " + os.path.join(directory,filename + ".pdf") + " created!",anchor="w",font="Helvetica 12 bold").grid(row = 0, column = 0)
            okButton = tk.Button(outDialog, text = "OK",command = outDialog.destroy)
            okButton.grid(row = 2, column = 2, padx = 5, pady = 5)
            for key in acc_keys:
                outDialog.bind(key, lambda event: outDialog.destroy())

    #ask for filename of report and ask for directory to store it in
    askDialog = tk.Toplevel(root)
    askDialog.geometry("550x100")
    askDialog.wm_title("Filename for Date Report")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    askDialog.iconphoto(False, ph1)
    tk.Label(askDialog,text="Name of report (no extension):",font="Helvetica 12").grid(row = 0, column = 0)
    fnd = tk.Entry(askDialog, width=25,font=("Helvetica", 15))
    fnd.grid(row = 0, column = 1, padx = 5, pady = 5)
    okButton = tk.Button(askDialog, text = "Accept", command=getFN)
    okButton.grid(row = 3, column = 1, padx = 5, pady = 5)
    cancelButton = tk.Button(askDialog, text = "Cancel", command=askDialog.destroy)
    cancelButton.grid(row = 3, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        fnd.bind(key, lambda event: getFN())
    fnd.focus_set() 
    
    



# button action definitions

#reports
def daterange():
    #return all entries with dates between the dates specified by the user
    def daterangeEvent():
        #pull dates from entry field
        strtdate = date_clean(start_date.get())
        enddate = date_clean(end_date.get())

        #do not proceed if both dates are 2000-01-01
        if (strtdate != '2000-01-01') or (enddate != '2000-01-01'):

            #make sure dates are in proper order
            if date_comp(strtdate, enddate) == 1:
                tmp = strtdate
                strtdate = enddate
                enddate = tmp
            daterangeDialog.destroy()
        
            #pull all entries from database between the dates given
            curs=cnx.cursor()
            curs.execute("SELECT * FROM Appeal_Records WHERE Date BETWEEN '" + str(strtdate) + "' AND '" + str(enddate) + "' ORDER BY Date, Time")
            data = curs.fetchall()
            #generate report
            if len(data) > 0:
                dateReport(fancy_date(strtdate),fancy_date(enddate),data)
            else:
                warnDialog = tk.Toplevel()
                warnDialog.grab_set()
                warnDialog.wm_title("ASC Database Notification")
                #application icon
                ph1 = tk.PhotoImage(file = icon_path)
                warnDialog.iconphoto(False, ph1)
                tk.Label(warnDialog,text="No entries found between " + fancy_date(strtdate) + " and " + fancy_date(enddate) + "!",font="Helvetica 12 bold").grid(row = 0, column = 0)
                closeButton = tk.Button(warnDialog, text = "OK",command = warnDialog.destroy)
                closeButton.grid(row = 3, column = 2, padx = 5, pady = 5)
                for key in acc_keys:
                    warnDialog.bind(key, lambda event: warnDialog.destroy())
        
    
    daterangeDialog = tk.Toplevel(root)
    daterangeDialog.grab_set()
    daterangeDialog.geometry("500x150")
    daterangeDialog.wm_title("Generate Report by Date Range")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    daterangeDialog.iconphoto(False, ph1)
    tk.Label(daterangeDialog,text="Enter all dates as YYYY-MM-DD.",font="Helvetica 12 bold").grid(row = 0, column = 3)
    tk.Label(daterangeDialog,text="Enter Starting Date:").grid(row = 3, column = 0)
    start_date=tk.Entry(daterangeDialog, width=25)
    start_date.grid(row = 3, column = 3, padx = 5, pady = 5)
    tk.Label(daterangeDialog,text="Enter Ending Date:").grid(row = 6, column = 0)
    end_date=tk.Entry(daterangeDialog, width=25)
    end_date.grid(row = 6, column = 3, padx = 5, pady = 5)
    gen_dr_reportButton = tk.Button(daterangeDialog, text = "Generate Report",command = daterangeEvent)
    gen_dr_reportButton.grid(row = 9, column = 3, padx = 5, pady = 5)
    cancelButton = tk.Button(daterangeDialog, text = "Cancel",command = daterangeDialog.destroy)
    cancelButton.grid(row = 9, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        daterangeDialog.bind(key, lambda event: daterangeEvent())
    start_date.focus()
    

def winreport():
    #return all entries with WINs matching those specified by the user
    def winreportEvent():
        #handle win report button event
        curs = cnx.cursor()

        #get WINs
        dat = win_entry.get('1.0', tk.END)
        dat = dat.strip()
        dat = dat.split(",")
        dat1=[]
        for x in dat:
            try:
                y=int(x)
                if y > 900000000 and y not in dat1:
                    dat1.append(y)
            except:
                pass
            
        #close dialog
        winreportDialog.destroy()
        
        #get DB entries for specified WINs
        results = []
        for win in dat1:
            cmd = "SELECT * FROM Appeal_Records WHERE WIN = "+str(win)
            curs.execute(cmd)
            fetch = curs.fetchall()
            if len(fetch) > 0:
                res=[win,fetch]
                results.append(res)
                
        #generate report
        if len(results) > 0:
            winReport(results)
        else:
            warnDialog = tk.Toplevel()
            warnDialog.grab_set()
            warnDialog.wm_title("ASC Database Notification")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            warnDialog.iconphoto(False, ph1)
            tk.Label(warnDialog,text="No entries found with the specified WINs!",font="Helvetica 12 bold").grid(row = 0, column = 0)
            closeButton = tk.Button(warnDialog, text = "OK",command = warnDialog.destroy)
            closeButton.grid(row = 3, column = 2, padx = 5, pady = 5)
            for key in acc_keys:
                warnDialog.bind(key, lambda event: warnDialog.destroy())


    winreportDialog = tk.Toplevel(root)
    winreportDialog.grab_set()
    winreportDialog.geometry("375x500")
    winreportDialog.wm_title("Generate Report from Specific WINs")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    winreportDialog.iconphoto(False, ph1)
    tk.Label(winreportDialog,text="Enter all WINs below.",font="Helvetica 12 bold").grid(row = 0, column = 0)
    tk.Label(winreportDialog,text="Format: 900123456 (no dashes) separated by commas").grid(row = 1, column = 0)
    win_entry = scrolledtext.ScrolledText(winreportDialog, wrap=tk.WORD, width=10, height=15,font=("Helvetica", 15))
    win_entry.grid(column=0, row=3, pady=10, padx=10)
    gen_dr_reportButton = tk.Button(winreportDialog, text = "Generate Report",command = winreportEvent)
    gen_dr_reportButton.grid(row = 5, column = 0, padx = 5, pady = 5)
    cancelButton = tk.Button(winreportDialog, text = "Cancel",command = winreportDialog.destroy)
    cancelButton.grid(row = 6, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        winreportDialog.bind(key, lambda event: winreportEvent())
    win_entry.focus()

#pull all entries corresponding to a single WIN; display each entry in tab of a notebook
def viewWIN():

    def viewEvent():
        curs = cnx.cursor()

        #get win then destroy dialog
        try:
            win = int(win_entry.get())
        except:
            win = 900000000
        if (win < 900000000) or (win > 999999999):
            win = 900000000
        askDialog.destroy()
        
        #pull entries matching given WIN
        curs.execute("SELECT * FROM Appeal_Records WHERE WIN = '" + str(win) + "' ORDER BY Date, Time")
        data = curs.fetchall()
        
        #if data, show it, else give notification
        if len(data) > 0:
            #create window to hold notebook
            notebookWindow = tk.Toplevel(root)
            notebookWindow.geometry("800x800")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            notebookWindow.iconphoto(False, ph1)
            notebookWindow.grab_set()
            notebookWindow.wm_title("DB Entries for " + data[0][1] + " " + data[0][2] + " (" + fancy_WIN(data[0][0]) + ")")

            #create notebook
            notebook = ttk.Notebook(notebookWindow)
            notebook.pack(pady=10, expand = True)
            

            #create a tab for each entry and pack into notebook
            tabs=[]
            for entry in data:
                tabs.append(ttk.Frame(notebook,width=800,height=1000))
            for i in range(0,len(data)):
                label1 = ttk.Label(tabs[i],text = "Committee Members:", font = "Helvetica 12 bold")
                label1.grid(row = 0, column = 0, sticky = "W")
                cm = ttk.Label(tabs[i], text = data[i][5], wraplength = "20c")
                cm.grid(row = 1, column = 0, sticky="W")
                label2 = ttk.Label(tabs[i], text = "Note Taker:",font = "Helvetica 12 bold")
                label2.grid(row = 2, column = 0, sticky = "W")
                nt = ttk.Label(tabs[i], text = data[i][6])
                nt.grid(row = 3, column = 0, sticky = "W")
                label3 = ttk.Label(tabs[i],text = " ")
                label3.grid(row = 4, column = 0)
                label4 = ttk.Label(tabs[i],text = "Minutes:", font = "Helvetica 12 bold")
                label4.grid(row = 5, column = 0, sticky = "W")
                mins = tk.Text(tabs[i], width = 85, height = 20, wrap='word')
                mins.insert(tk.INSERT,data[i][7])
                mins.configure(state='disabled')
                mins.grid(row = 6, column = 0)
                vsb = ttk.Scrollbar(tabs[i], command=mins.yview, orient="vertical")
                mins.configure(yscrollcommand=vsb.set)
                vsb.grid(row=6, column=1, sticky="ns")
                label5 = ttk.Label(tabs[i],text = "Motion:", font = "Helvetica 12 bold")
                label5.grid(row = 7, column = 0, sticky = "W")
                motion = ttk.Label(tabs[i], text = data[i][8], wraplength = "20c")
                motion.grid(row = 8, column = 0, sticky="W")
                label6 = ttk.Label(tabs[i],text = "Decision:", font = "Helvetica 12 bold")
                label6.grid(row = 9, column = 0, sticky = "W")
                dec = ttk.Label(tabs[i], text = data[i][9], wraplength = "20c")
                dec.grid(row = 10, column = 0, sticky="W")
                spc = ttk.Label(tabs[i],text = " ")
                spc.grid(row = 11, column = 0, sticky = "W")
                label7 = ttk.Label(tabs[i],text = "Readmit: " + data[i][10], font = "Helvetica 12 bold")
                label7.grid(row = 12, column = 0, sticky = "W")
                tabs[i].pack(fill='both', expand=True)
                notebook.add(tabs[i], text = fancy_date(str(data[i][3])) + " at " + data[i][4] )
            closeButton = tk.Button(notebookWindow, text = "Close", command = notebookWindow.destroy)
            closeButton.pack(padx = 5, pady = 5)
            for key in acc_keys:
                notebookWindow.bind(key, lambda event: notebookWindow.destroy())
        else:
            nodataDialog = tk.Toplevel(root)
            nodataDialog.geometry("400x75")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            nodataDialog.iconphoto(False, ph1)
            nodataDialog.grab_set()
            nodataDialog.wm_title("Single WIN Search")
            tk.Label(nodataDialog,text="There are no entries matching WIN = " + fancy_WIN(win) + "!", font="Helvetica 12 bold").grid(row = 0, column = 0)
            okButton = tk.Button(nodataDialog, text = "OK", command = nodataDialog.destroy)
            okButton.grid(row = 1, column = 0, padx = 5, pady = 5)
            for key in acc_keys:
                nodataDialog.bind(key, lambda event: nodataDialog.destroy())
        
    #ask for WIN
    askDialog = tk.Toplevel(root)
    askDialog.geometry("500x100")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    askDialog.iconphoto(False, ph1)
    askDialog.grab_set()
    askDialog.wm_title("Single WIN Search")
    tk.Label(askDialog,text="Enter WIN in the format 900123456 (no dashes): ",font="Helvetica 12 bold").grid(row = 0, column = 0)
    win_entry = tk.Entry(askDialog, width=10,font=("Helvetica", 15))
    win_entry.grid(column=2, row=0, pady=10, padx=10)
    win_entry.focus()
    gen_reportButton = tk.Button(askDialog, text = "View",command = viewEvent)
    gen_reportButton.grid(row = 1, column = 2, padx = 5, pady = 5)
    cancelButton = tk.Button(askDialog, text = "Cancel", command = askDialog.destroy)
    cancelButton.grid(row = 1, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        askDialog.bind(key, lambda event: viewEvent())
    
def dbstats():
    #pulls important stats about the Appeals DB
    curs=cnx.cursor()
    
    def detailedStats(tot, early, last, succ, deny):
        #generates a detailed report about the Appeals_DB and saves it to a PDF file.
        succpc = round(succ/tot*100,2)
        denypc = round(deny/tot*100,2)

        #create list to count number of successful and denied appeals broken down by first-time, second-time, third-time, and higher-time appeals
        cnts = [[0,0], [0,0],[0,0],[0,0]]
        #get all distinct WINs
        curs.execute("SELECT DISTINCT WIN FROM Appeal_Records")
        wins = curs.fetchall()
        dist = len(wins)
        
        #cycle through wins, for each one update cnts
        for win in wins:
            curs.execute("SELECT * FROM Appeal_Records WHERE WIN = '" + str(win[0]) + "' ORDER BY Date, Time")
            entries = curs.fetchall()
            #pop first entry
            entry = entries.pop(0)
            if entry[10] == 'Y':
                cnts[0][0]+=1
            else:
                cnts[0][1]+=1

            #check to see if a second entry exists
            if len(entries) > 0:
                entry = entries.pop(0)
                if entry[10] == 'Y':
                    cnts[1][0]+=1
                else:
                    cnts[1][1]+=1

            #check to see if a third entry exists
            if len(entries) > 0:
                entry = entries.pop(0)
                if entry[10] == 'Y':
                    cnts[2][0]+=1
                else:
                    cnts[2][1]+=1

            #check to see if there are more than 3 entries
            if len(entries) > 0:
                for entry in entries:
                    if entry[10] == 'Y':
                        cnts[3][0]+=1
                    else:
                        cnts[3][1]+=1

        t1 = cnts[0][0] + cnts[0][1]
        t1pc = round(t1/tot*100,2)
        t2 = cnts[1][0] + cnts[1][1]
        t2pc = round(t2/tot*100,2)
        t3 = cnts[2][0] + cnts[2][1]
        t3pc = round(t3/tot*100,2)
        t4 = cnts[3][0] + cnts[3][1]
        t4pc = round(t4/tot*100,2)
        
        #generate standard stats file name
        now = datetime.now()
        filename = now.strftime("%Y_%m_%d_%H%M")
        filename = "AppealsDB_Stats_"+filename

        #ask for directory to save file to
        statsDialog.destroy()
        direc = fd.askdirectory(title= "Please select directory for report:")
        if direc == () or direc == '':
                return 1
        filename = os.path.join(direc,filename)

        #open tex file for writing
        file = open(filename+".tex","w")
        
        #write coverpage
        for line in stdhd:
            file.write(line)
        file.write("\\title{Academic Standards Committee\\\\{\\huge \\textbf{DETAILED REPORT on the \\\\ APPEALS DATABASE}} \\vspace{-20pt}}\n")
        file.write("\\begin{document}\n")
        file.write("\\maketitle\n")
        file.write("\\thispagestyle{empty}\n")
        file.write("\\begin{tabular}{rl}\n")
        file.write("\\noindent {\\Large \\textbf{Total Number of Entries: } } & {\\Large " + str(tot) + "} \\\\ \n")
        file.write(" \n")
        file.write("\\noindent {\\Large \\textbf{Number of Distinct WINs: } } & {\\Large " + str(dist) + "} \\\\ \n")
        file.write(" \n")
        file.write(" & \\\\ \n")
        file.write("\\noindent {\\Large \\textbf{Earliest Entry: } } & {\\Large " + early + "} \\\\ \n")
        file.write("\\noindent {\\Large \\textbf{Latest Entry: } } & {\\Large " + last + "} \\\\ \n")
        file.write(" & \\\\ \n")
        file.write("\\noindent {\\Large \\textbf{Successful Appeals: } } & {\\Large " + str(succ) + " (" + str(succpc) + "\\%)} \\\\ \n")
        file.write("\\noindent {\\Large \\textbf{Denied Appeals: } } & {\\Large " + str(deny) + " (" + str(denypc) + "\\%)} \\\\ \n")
        file.write("\\end{tabular} \n")
        file.write(" \n")
        file.write(" \\bigskip \n")
        file.write(" \n")
        file.write("\\noindent {\\Large \\textbf{Database Details:}} \n")
        file.write(" \\begin{itemize}\n")
        file.write(" \\item \\textbf{First-Time Appeals:}\n")
        file.write(" \n")
        file.write(" Number of First-Time Appeals: " + str(t1) + " (" + str(t1pc) + "\\% of total appeals)" + "\\\\ \n")
        file.write(" Number Approved: " + str(cnts[0][0]) + " (" + str(round(cnts[0][0]/t1*100,2)) + "\\% of first-time appeals)" + "\\\\ \n")
        file.write(" Number Denied: " + str(cnts[0][1]) + " (" + str(round(cnts[0][1]/t1*100,2)) + "\\% of first-time appeals)" + " \n")
        file.write(" \\item \\textbf{Second-Time Appeals:}\n")
        file.write(" \n")
        file.write(" Number of Second-Time Appeals: " + str(t2) + " (" + str(t2pc) + "\\% of total appeals)" + " \\\\ \n")
        file.write(" Number Approved: " + str(cnts[1][0]) + " (" + str(round(cnts[1][0]/t2*100,2)) + "\\% of second-time appeals)" + "\\\\ \n")
        file.write(" Number Denied: " + str(cnts[1][1]) + " (" + str(round(cnts[1][1]/t2*100,2)) + "\\% of second-time appeals)" + " \n")
        file.write(" \\item \\textbf{Third-Time Appeals:}\n")
        file.write(" \n")
        file.write(" Number of Third-Time Appeals: " + str(t3) + " (" + str(t3pc) + "\\% of total appeals)" + "\\\\ \n")
        file.write(" Number Approved: " + str(cnts[2][0]) + " (" + str(round(cnts[2][0]/t3*100,2)) + "\\% of third-time appeals)" + "\\\\ \n")
        file.write(" Number Denied: " + str(cnts[2][1]) + " (" + str(round(cnts[2][1]/t3*100,2)) + "\\% of third-time appeals)" + " \n")
        file.write(" \\item \\textbf{Fourth-Time and Higher Appeals:}\n")
        file.write(" \n")
        file.write(" Number of Appeals Fourth-Time and Higher: " + str(t4) + " (" + str(t4pc) + "\\% of total appeals)" + "\\\\ \n")
        file.write(" Number Approved: " + str(cnts[3][0]) + " (" + str(round(cnts[3][0]/t4*100,2)) + "\\% of fourth-time and higher appeals)" + "\\\\ \n")
        file.write(" Number Denied: " + str(cnts[3][1]) + " (" + str(round(cnts[3][1]/t4*100,2)) + "\\% of fourth-time and higher appeals)" + " \n")
        file.write(" \\end{itemize}\n")
        file.write(" \n")
        file.write("\\noindent {\\footnotesize \\textbf{Warning:} Entries from the earliest appeal dates may be incorrectly counted as first-time, second-time, etc. as any previous appearances before ASC are not recorded in the database!} \n")
        file.write("\\end{document}\n")
        file.close()
        
        #write shell script to compile tex to pdf and erase auxiliary files
        shscr = os.path.join(direc,"tmp.sh")
        file = open(shscr,"w")
        file.write("#!/bin/sh \n")
        file.write("/usr/bin/xelatex -output-directory " + direc + " " + filename + ".tex" +"\n")
        file.write("/usr/bin/xelatex -output-directory " + direc + " " + filename + ".tex" +"\n")
        file.write("rm " + filename + ".tex" + " \n")
        file.write("rm " +filename + ".log " + "\n")
        file.write("rm " + filename + ".aux " + "\n")
        file.write("rm " + filename + ".out " + "\n")
        file.write("rm " + os.path.join(direc,"texput.log") + "\n")
        file.write("rm " + shscr)
        file.close()
        
        #run shell script
        os.system("sh " + shscr)

        #success dialog!
        outDialog = tk.Toplevel(root)
        outDialog.wm_title("Detailed Stats Report")
        #application icon
        ph1 = tk.PhotoImage(file = icon_path)
        outDialog.iconphoto(False, ph1)
        outDialog.grab_set()
        tk.Label(outDialog,text= "Detailed Stats Report: " + os.path.join(direc,filename + ".pdf") + " created!",anchor="w",font="Helvetica 12 bold").grid(row = 0, column = 0)
        okButton = tk.Button(outDialog, text = "OK",command = outDialog.destroy)
        okButton.grid(row = 2, column = 2, padx = 5, pady = 5)
        for key in acc_keys:
            outDialog.bind(key, lambda event: outDialog.destroy())

    try:
        curs.execute("SELECT COUNT(WIN) FROM Appeal_Records")
    except mysql.connector.Error as err:
        tot = 0
    else:
        tot=curs.fetchall()[0][0]

    #get earliest appeal date
    try:
        curs.execute("SELECT MIN(Date) FROM Appeal_Records")
    except mysql.connector.Error as err:
        mindate = "ERROR"
    else:
        dt = str(curs.fetchall()[0][0])
        if dt == "None":
            mindate = "ERROR"
        else:
            mindate=fancy_date(dt)
    

    #get lastest appeal date
    try:
        curs.execute("SELECT MAX(Date) FROM Appeal_Records")
    except mysql.connector.Error as err:
        maxdate = "ERROR"
    else:
        dt = str(curs.fetchall()[0][0])
        if dt == "None":
            maxdate = "ERROR"
        else:
            maxdate=fancy_date(dt)

    #get number of distinct WINs
    try:
        curs.execute("SELECT COUNT(DISTINCT WIN) FROM Appeal_Records")
    except mysql.connector.Error as err:
        wins = 0
    else:
        wins=curs.fetchall()[0][0]

    #get number of successful appeals
    try:
        curs.execute("SELECT COUNT(WIN) FROM Appeal_Records WHERE Readmit=\"Y\"")
    except mysql.connector.Error as err:
        succ=0
    else:
        succ=curs.fetchall()[0][0]
    
    #get number of denied appeals
    try:
        curs.execute("SELECT COUNT(WIN) FROM Appeal_Records WHERE Readmit=\"N\"")
    except mysql.connector.Error as err:
        deny = 0
    else:
        deny=curs.fetchall()[0][0]


    statsDialog = tk.Toplevel(root)
    statsDialog.grab_set()
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    statsDialog.iconphoto(False, ph1)
    statsDialog.geometry("375x250")
    statsDialog.wm_title("ASC Database Statistics")
    tk.Label(statsDialog,text="").grid(row = 0, column = 0, sticky="e")
    tk.Label(statsDialog,text="Total number of records: ", font="Helvetica 12 bold").grid(row = 2, column = 0, sticky="e")
    tk.Label(statsDialog,text=str(tot), font="Helvetica 12").grid(row = 2, column = 1, sticky="w")
    tk.Label(statsDialog,text="Earliest Appeal: ", font="Helvetica 12 bold").grid(row = 4, column = 0, sticky="e")
    tk.Label(statsDialog,text=str(mindate), font="Helvetica 12").grid(row = 4, column = 1, sticky="w")
    tk.Label(statsDialog,text="Latest Appeal: ", font="Helvetica 12 bold").grid(row = 6, column = 0, sticky="e")
    tk.Label(statsDialog,text=str(maxdate), font="Helvetica 12").grid(row = 6, column = 1, sticky="w")
    tk.Label(statsDialog,text="").grid(row = 8, column = 0, sticky="e")
    tk.Label(statsDialog,text="Distinct WINs: ", font="Helvetica 12 bold").grid(row = 10, column = 0, sticky="e")
    tk.Label(statsDialog,text=str(wins), font="Helvetica 12").grid(row = 10, column = 1, sticky="w")
    tk.Label(statsDialog,text="Successful appeals: ", font="Helvetica 12 bold").grid(row = 12, column = 0,sticky="e")
    tk.Label(statsDialog,text=str(succ), font="Helvetica 12").grid(row = 12, column = 1, sticky="w")
    tk.Label(statsDialog,text="Denied appeals: ", font="Helvetica 12 bold").grid(row = 14, column = 0, sticky="e")
    tk.Label(statsDialog,text=str(deny), font="Helvetica 12").grid(row = 14, column = 1, sticky="w")
    tk.Label(statsDialog,text="").grid(row = 16, column = 0, sticky="e")
    closeButton = tk.Button(statsDialog, text = "Close",command = statsDialog.destroy)
    closeButton.grid(row = 18, column = 1,sticky="w")
    if wins > 0:
        detailButton = tk.Button(statsDialog, text = "Detailed Stats Report",command = partial(detailedStats, tot, mindate, maxdate, succ, deny))
        detailButton.grid(row = 18, column = 0,sticky="w")
    for key in acc_keys:
        statsDialog.bind(key, lambda event: statsDialog.destroy())

#management
def add1():
    #add a single entry to db
    createEntry()

def addfl():
    #function defining add from file action
    def addAction():
        warnDialog.destroy()
        curs = cnx.cursor()
        #add multiple entries from a csv file
        file = fd.askopenfile(mode = 'r',title= "Please select file:",filetypes= (('CSV files', '*.csv'),))
        if file == '' or file == ():
            return 1
    
        entries_reader = csv.reader(file, delimiter='\t')
        errors = []
        tot = 0
        for datapt in entries_reader:
            if len(datapt) == 11:
                entry = dbEntry()
                entry.win = int(datapt[0])
                entry.date = date_clean(str(datapt[3]))
                entry.time = str(datapt[4])
                entry.first_name = datapt[1]
                entry.last_name = datapt[2]
                entry.comm = datapt[5]
                entry.nt = datapt[6]
                entry.mins = datapt[7]
                entry.motion = datapt[8]
                entry.dec = datapt[9]
                try:
                    entry.readmit = datapt[10].strip(' ')[0]
                except:
                    entry.readmit = '?'
                verify = verifyFields(entry)
                #verify necessary fields, collect bad entries in error list with error messages
                if (entry.date.split("-")[0] == '2000') or (verify.win == 1) or (verify.first == 1) or (verify.last == 1) or (verify.mins == 1) or (verify.readmit == 1):
                    msg = [[str(entry.win), entry.first_name, entry.last_name, entry.date, entry.time]]
                    if verify.win == 1 :
                        msg.append("ASC_DB Error: WIN incorrectly formatted.")
                    if verify.first == 1 or verify.last == 1:
                        msg.append("ASC_DB Error: Student first and/or last name is missing.")
                    if verify.mins == 1:
                        msg.append("ASC_DB Error: No minutes recorded.")
                    if verify.readmit == 1:
                        msg.append("ASC_DB Error: Readmit field incorrectly formatted.")
                    if entry.date.split("-")[0] == '2000':
                        msg.append("ASC_DB Error: Suspicious date.")
                    errors.append(msg)
                else:
                    #add correct entries to DB
                    cmd = "INSERT IGNORE INTO Appeal_Records (WIN, First_Name, Last_Name, Date, Time, Committee_Members, Note_Taker, Minutes, Motion, Decision, Readmit) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                    try:
                        curs.execute(cmd,(str(entry.win), entry.first_name, entry.last_name, entry.date, entry.time, entry.comm, entry.nt, entry.mins, entry.motion, entry.dec, entry.readmit))
                        cnx.commit()
                        tot+=1
                    except mysql.connector.Error as bad:
                        if bad.errno == 1265:
                            msg = [[str(entry.win), entry.first_name, entry.last_name, str(entry.date), entry.time], 'MySQL Error: ' + str(bad)]
                            errors.append(msg)
                            cmd =  "DELETE FROM Appeal_Records WHERE WIN = %s AND Date = %s"
                            curs.execute(cmd, (str(entry.win), str(entry.date)))
                            cnx.commit()
                        else:
                            msg = [[str(entry.win), entry.first_name, entry.last_name, str(entry.date), entry.time], 'MySQL Error: ' + str(bad)]
                            errors.append(msg)
                    
            else:
                #collect entries with too few fields
                dat=[datapt[0]]
                try:
                    dat.append(datapt[1])
                except:
                    dat.append(" ")
                try:
                    dat.append(datapt[2])
                except:
                    dat.append(" ")
                try:
                    dat.append(datapt[3])
                except:
                    dat.append(" ")
                try:
                    dat.append(datapt[4])
                except:
                    dat.append(" ")
                errors.append([dat,"ASC_DB Error: Number of fields less than expected."])
                
        #error reporting function
        def errorReport():
            #create window to hold notebook
            notebookWindow = tk.Toplevel(root)
            notebookWindow.geometry("1000x400")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            notebookWindow.iconphoto(False, ph1)
            notebookWindow.grab_set()
            notebookWindow.wm_title("Add Entries from File: Error Details")
            #create notebook
            notebook = ttk.Notebook(notebookWindow)
            notebook.pack(pady=10, expand = True)

            #create a tab for each entry and pack into notebook
            tabs=[]
            for i in range(len(errors)):
                tabs.append(ttk.Frame(notebook,width=1000,height=400))
                label1 = ttk.Label(tabs[i],text = "WIN: ", font = "Helvetica 12 bold")
                label1.grid(row = 0, column = 0, sticky = "E")
                wn = ttk.Label(tabs[i], text = errors[i][0][0])
                wn.grid(row = 0, column = 1, sticky="W")
                label2 = ttk.Label(tabs[i], text = "First Name: ",font = "Helvetica 12 bold")
                label2.grid(row = 1, column = 0, sticky = "E")
                fn = ttk.Label(tabs[i], text = errors[i][0][1])
                fn.grid(row = 1, column = 1, sticky = "W")
                label3 = ttk.Label(tabs[i],text = "Last Name: ", font = "Helvetica 12 bold")
                label3.grid(row = 2, column = 0,sticky = "E")
                ln = ttk.Label(tabs[i], text = errors[i][0][2])
                ln.grid(row = 2, column = 1, sticky = "W")
                label5 = ttk.Label(tabs[i],text = "Date: ", font = "Helvetica 12 bold")
                label5.grid(row = 3, column = 0, sticky = "E")
                date = ttk.Label(tabs[i], text = errors[i][0][3])
                date.grid(row = 3, column = 1, sticky="W")
                label6 = ttk.Label(tabs[i],text = "Time: ", font = "Helvetica 12 bold")
                label6.grid(row = 4, column = 0, sticky = "E")
                time = ttk.Label(tabs[i], text = errors[i][0][4])
                time.grid(row = 4, column = 1, sticky="W")
                spc = ttk.Label(tabs[i],text = " ")
                spc.grid(row = 5, column = 0, sticky = "W")
                label7 = ttk.Label(tabs[i],text = "Errors:", font = "Helvetica 12 bold")
                label7.grid(row = 6, column = 0, sticky = "E")
                for j in range(1,len(errors[i])):
                    err = ttk.Label(tabs[i],text = errors[i][j])
                    err.grid(row = 6 + j, column = 1, sticky="W")
                tabs[i].pack(fill='both', expand=True)
                notebook.add(tabs[i], text =  str(i+1))
            closeButton = tk.Button(notebookWindow, text = 'Close', command = notebookWindow.destroy)
            closeButton.pack()
            for key in acc_keys:
                notebookWindow.bind(key, lambda event: notebookWindow.destroy()) 
            
        #result dialog
        outcomeDialog = tk.Toplevel(root)
        outcomeDialog.geometry("400x150")
        #application icon
        ph1 = tk.PhotoImage(file = icon_path)
        outcomeDialog.iconphoto(False, ph1)
        outcomeDialog.grab_set()
        outcomeDialog.wm_title("Add Entries from File")
        tk.Label(outcomeDialog,text="Disposition of Entries Added from File:", font="Helvetica 12 bold").grid(row = 0, column = 0, sticky="w")
        tk.Label(outcomeDialog,text="Total Entries Added: " + str(tot), font="Helvetica 12").grid(row = 1, column = 0, sticky="w")
        tk.Label(outcomeDialog,text="Entries with Errors: " + str(len(errors)), font="Helvetica 12").grid(row = 2, column = 0, sticky="w")
        tk.Label(outcomeDialog,text=" ", font="Helvetica 12").grid(row = 3, column = 0, sticky="w")
        closeButton = tk.Button(outcomeDialog, text = "Close",command = outcomeDialog.destroy)
        closeButton.grid(row = 4, column = 0,sticky="w")
        for key in acc_keys:
            outcomeDialog.bind(key, lambda event: outcomeDialog.destroy()) 
        if len(errors) > 0:
            errorsButton = tk.Button(outcomeDialog, text = "View Errors",command = errorReport)
            errorsButton.grid(row = 4, column = 1,sticky="w")
            
        
      
        
        

    #warning dialog   
    warnDialog = tk.Toplevel(root)
    warnDialog.geometry("950x200")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    warnDialog.iconphoto(False, ph1)
    warnDialog.grab_set()
    warnDialog.wm_title("Add Entries from File")
    tk.Label(warnDialog,text="WARNING: This procedure requires a CSV input file which must be formatted STRICTLY AS FOLLOWS.", font="Helvetica 12 bold").grid(row = 0, column = 0, sticky="w")
    tk.Label(warnDialog,text="- Fields in each entry should be delimited by TABS.", font="Helvetica 12").grid(row = 1, column = 0, sticky="w")
    tk.Label(warnDialog,text="- Individual entries should be separated by NEWLINES.", font="Helvetica 12").grid(row = 2, column = 0, sticky="w")
    tk.Label(warnDialog,text="- Field Order: WIN, First_Name, Last_Name, Date, Time, Committee_Members, Note_Taker, Minutes, Motion, Decision, Readmit?", font="Helvetica 12").grid(row = 3, column = 0, sticky="w")
    tk.Label(warnDialog,text="- MAKE SURE ALL DATA IS CORRECT BEFORE PROCEEDING!", font="Helvetica 12").grid(row = 4, column = 0, sticky="w")
    tk.Label(warnDialog,text=" ", font="Helvetica 12").grid(row = 5, column = 0, sticky="w")
    closeButton = tk.Button(warnDialog, text = "Cancel",command = warnDialog.destroy)
    closeButton.grid(row = 6, column = 0,sticky="e")
    proceedButton = tk.Button(warnDialog, text = "Proceed",command = addAction)
    proceedButton.grid(row = 6, column = 0,sticky="w")
    for key in acc_keys:
        warnDialog.bind(key, lambda event: addAction())

#modify an existing entry
def modify():
    curs = cnx.cursor()

    #modify event function
    #pull dates associated to WIN
    #if one date only - go to modification screen
    #if multiple dates - ask user to specify which entry, pull entry, go to modification screen
    def modifyEvent():
        WIN = win.get()
        if (int(WIN) > 900000000) & (int(WIN) < 900999999):
            modifyDialog.destroy()
            curs.execute("SELECT * FROM Appeal_Records WHERE WIN = '" + WIN + "'")
            entries = curs.fetchall()
            def entryCapture(i):
                entry = dbEntry()
                entry.win = int(entries[i][0])
                entry.date = str(entries[i][3])
                entry.time = str(entries[i][4])
                entry.first_name = entries[i][1]
                entry.last_name = entries[i][2]
                entry.comm = entries[i][5]
                entry.nt = entries[i][6]
                entry.mins = entries[i][7]
                entry.motion = entries[i][8]
                entry.dec = entries[i][9]
                entry.readmit = entries[i][10]
                verifyEntry(entry)
            
            if len(entries) > 1:
                chs = tk.IntVar()
                choiceDialog = tk.Toplevel(root)
                choiceDialog.geometry("400x200")
                #application icon
                ph1 = tk.PhotoImage(file = icon_path)
                choiceDialog.iconphoto(False, ph1)
                choiceDialog.grab_set()
                choiceDialog.wm_title("Existing Entries for WIN = " + fancy_WIN(int(WIN)))
                tk.Label(choiceDialog,text="Select 1 of the following: ", font="Helvetica 12 bold").grid(row = 0, column = 0)
                for i in range(len(entries)):
                    R = tk.Radiobutton(choiceDialog, text=fancy_date(str(entries[i][3])), variable=chs, value=i)
                    R.grid(row = i+1, column = 0 )
                okButton = tk.Button(choiceDialog, text = "Modify", command = lambda:[entryCapture(int(str(chs.get()))), choiceDialog.destroy()])#start editing process on chosen entry
                okButton.grid(row = len(entries)+1, column = 1, padx = 5, pady = 5)
            
            elif len(entries) == 1:
                entryCapture(0)
        
            else:
                nodataDialog = tk.Toplevel(root)
                nodataDialog.geometry("400x75")
                #application icon
                ph1 = tk.PhotoImage(file = icon_path)
                nodataDialog.iconphoto(False, ph1)
                nodataDialog.grab_set()
                nodataDialog.wm_title("Modify Existing Entry")
                tk.Label(nodataDialog,text="There are no entries matching WIN = " + fancy_WIN(int(WIN)) + "!", font="Helvetica 12 bold").grid(row = 0, column = 0)
                okButton = tk.Button(nodataDialog, text = "OK", command = nodataDialog.destroy)
                okButton.grid(row = 1, column = 0, padx = 5, pady = 5)
            
    
    #ask for WIN
    modifyDialog = tk.Toplevel(root)
    modifyDialog.grab_set()
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    modifyDialog.iconphoto(False, ph1)
    modifyDialog.geometry("400x100")
    modifyDialog.wm_title("Modify Existing Entry")
    tk.Label(modifyDialog,text="Enter WIN of DB entry to be modified:").grid(row = 0, column = 0)
    win=tk.Entry(modifyDialog, width=10,font=("Helvetica", 15))
    win.grid(row = 0, column = 1, padx = 5, pady = 5)
    modifyButton = tk.Button(modifyDialog, text = "Submit",command = modifyEvent)
    modifyButton.grid(row = 2, column = 1, padx = 5, pady = 5)
    cancelButton = tk.Button(modifyDialog, text = "Cancel",command = modifyDialog.destroy)
    cancelButton.grid(row = 2, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        modifyDialog.bind(key, lambda event: modifyEvent())
    win.focus()
    

def deletewin():
    #delete entries matching given WINs
    def windeleteEvent():
        #handle win delete button event
        dat = win_entry.get('1.0', tk.END)
        dat = dat.strip()
        dat = dat.split(",")
        dat1=[]
        for x in dat:
            try:
                y=int(x)
                if y > 900000000:
                    dat1.append(y)
            except:
                pass
        if len(dat1) > 0 :
            fancywins= "Selected WINs: "
            for x in dat1:
                fancywins = fancywins + str(x) + ", "
            fancywins = fancywins[:-2]

            def deleteAction():
                cnt = 0
                curs = cnx.cursor()
                for x in dat1:
                    cmd = "DELETE FROM Appeal_Records WHERE WIN = %s"
                    try:
                        curs.execute(cmd, (str(x),))
                        cnx.commit()
                    except:
                        pass
                    cnt += curs.rowcount
                    
                confirmDialog.destroy()
                outDialog = tk.Toplevel(root)
                outDialog.wm_title("ASC DB Delete")
                #application icon
                ph1 = tk.PhotoImage(file = icon_path)
                outDialog.iconphoto(False, ph1)
                outDialog.grab_set()
                tk.Label(outDialog,text= str(cnt) + " entries deleted from database.",anchor="w",font="Helvetica 12 bold").grid(row = 0, column = 0)
                okButton = tk.Button(outDialog, text = "OK",command = outDialog.destroy)
                okButton.grid(row = 2, column = 2, padx = 5, pady = 5)
                for key in acc_keys:
                    outDialog.bind(key, lambda event: outDialog.destroy())
               
       
            windeleteDialog.destroy()
            confirmDialog = tk.Toplevel(root)
            confirmDialog.grab_set()
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            confirmDialog.iconphoto(False, ph1)
            tk.Label(confirmDialog,text="WARNING: All DB entries matching the following WINs will be permanently removed.",font="Helvetica 12 bold",anchor="w").grid(row = 0, column = 0)
            tk.Label(confirmDialog,text=fancywins,font="Helvetica 12 bold",anchor="w").grid(row = 2, column = 0)
            tk.Label(confirmDialog,text="Are you sure you want to proceed?",anchor="w").grid(row = 4, column = 0)
            proceedButton = tk.Button(confirmDialog, text = "Proceed",command = deleteAction)
            proceedButton.grid(row = 6, column = 0, padx = 5, pady = 5)
            cancelButton = tk.Button(confirmDialog, text = "Cancel",command = confirmDialog.destroy)
            cancelButton.grid(row = 6, column = 2, padx = 5, pady = 5)
            for key in acc_keys:
                confirmDialog.bind(key, lambda event: deleteAction())

    windeleteDialog = tk.Toplevel(root)
    windeleteDialog.grab_set()
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    windeleteDialog.iconphoto(False, ph1)
    windeleteDialog.geometry("375x500")
    windeleteDialog.wm_title("Delete Entries with Specific WINs")
    tk.Label(windeleteDialog,text="Enter all WINs below.",font="Helvetica 12 bold").grid(row = 0, column = 0)
    tk.Label(windeleteDialog,text="Format: 900123456 (no dashes) separated by commas").grid(row = 1, column = 0)
    win_entry = scrolledtext.ScrolledText(windeleteDialog, wrap=tk.WORD, width=10, height=15,font=("Helvetica", 15))
    win_entry.grid(column=0, row=3, pady=10, padx=10)
    gen_dr_reportButton = tk.Button(windeleteDialog, text = "Proceed",command = windeleteEvent)
    gen_dr_reportButton.grid(row = 5, column = 0, padx = 5, pady = 5)
    cancelButton = tk.Button(windeleteDialog, text = "Cancel",command = windeleteDialog.destroy)
    cancelButton.grid(row = 6, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        windeleteDialog.bind(key, lambda event: windeleteEvent())
    win_entry.focus()

def deletedate():
    #delete entries between specified dates
    def daterangeEvent():
        def deleteAction():
            curs = cnx.cursor()
            cmd = "DELETE FROM Appeal_Records WHERE Date BETWEEN %s AND %s"
            curs.execute(cmd, (strtdate, enddate))
            cnx.commit()
            confirmDialog.destroy()
            outDialog = tk.Toplevel(root)
            outDialog.wm_title("ASC DB Delete")
            #application icon
            ph1 = tk.PhotoImage(file = icon_path)
            outDialog.iconphoto(False, ph1)
            outDialog.grab_set()
            tk.Label(outDialog,text= str(curs.rowcount) + " entries deleted from database.",anchor="w",font="Helvetica 12 bold").grid(row = 0, column = 0)
            okButton = tk.Button(outDialog, text = "OK",command = outDialog.destroy)
            okButton.grid(row = 2, column = 2, padx = 5, pady = 5)
            for key in acc_keys:
                outDialog.bind(key, lambda event: outDialog.destroy())
            
        strtdate = date_clean(start_date.get())
        enddate = date_clean(end_date.get())
        if date_comp(strtdate, enddate) == 1:
            tmp = strtdate
            strtdate = enddate
            enddate = tmp
        daterangeDialog.destroy()
        confirmDialog = tk.Toplevel(root)
        confirmDialog.grab_set()
        #application icon
        ph1 = tk.PhotoImage(file = icon_path)
        confirmDialog.iconphoto(False, ph1)
        tk.Label(confirmDialog,text="WARNING: All entries between "+ fancy_date(strtdate) + " and " + fancy_date(enddate) + " will be permanently removed.",anchor="w",font="Helvetica 12 bold").grid(row = 0, column = 0)
        tk.Label(confirmDialog,text="Are you sure you want to proceed?",anchor="w").grid(row = 2, column = 0)
        proceedButton = tk.Button(confirmDialog, text = "Proceed",command = deleteAction)
        proceedButton.grid(row = 4, column = 0, padx = 5, pady = 5)
        cancelButton = tk.Button(confirmDialog, text = "Cancel",command = confirmDialog.destroy)
        cancelButton.grid(row = 4, column = 2, padx = 5, pady = 5)
        for key in acc_keys:
            confirmDialog.bind(key, lambda event: deleteAction())
        

    daterangeDialog = tk.Toplevel(root)
    daterangeDialog.grab_set()
    daterangeDialog.geometry("450x150")
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    daterangeDialog.iconphoto(False, ph1)
    daterangeDialog.wm_title("Delete Entries by Date Range")
    tk.Label(daterangeDialog,text="Enter all dates as YYYY-MM-DD.",font="Helvetica 12 bold").grid(row = 0, column = 3)
    tk.Label(daterangeDialog,text="Enter Starting Date:").grid(row = 3, column = 0)
    start_date=tk.Entry(daterangeDialog, width=25)
    start_date.grid(row = 3, column = 3, padx = 5, pady = 5)
    tk.Label(daterangeDialog,text="Enter Ending Date:").grid(row = 6, column = 0)
    end_date=tk.Entry(daterangeDialog, width=25)
    end_date.grid(row = 6, column = 3, padx = 5, pady = 5)
    dr_deleteButton = tk.Button(daterangeDialog, text = "Delete Entries",command = daterangeEvent)
    dr_deleteButton.grid(row = 9, column = 3, padx = 5, pady = 5)
    cancelButton = tk.Button(daterangeDialog, text = "Cancel",command = daterangeDialog.destroy)
    cancelButton.grid(row = 9, column = 0, padx = 5, pady = 5)
    for key in acc_keys:
        daterangeDialog.bind(key, lambda event: daterangeEvent())
    start_date.focus()



#backup and restore
def backup():
    #backs up the current db

    #generate standard backup file name
    now = datetime.now()
    filename = now.strftime("%Y_%m_%d_%H%M")
    filename = "AppealsDB_BackUp_"+filename+".sql"

    #ask for directory to save file to
    direc = fd.askdirectory(title= "Please select directory for backup file:")
    if direc == () or direc == '':
                return 1

    #generate dump command
    cmd = "mysqldump -u root -p" + pswd + " Appeals_DB > " + direc + "/" + filename

    #try dump, notify of failure or success
    try:
        os.popen(cmd)
    except:
        backupDialog = tk.Toplevel(root)
        backupDialog.wm_title("ASC Database Backup Dialog")
        #application icon
        ph1 = tk.PhotoImage(file = icon_path)
        backupDialog.iconphoto(False, ph1)
        tk.Label(backupDialog,text="Backup Procedure Failed!",font="Helvetica 12 bold").grid(row = 0, column = 0)
        okButton = tk.Button(backupDialog, text = "OK", command=backupDialog.destroy)
        okButton.grid(row = 3, column = 1, padx = 5, pady = 5)
        for key in acc_keys:
            backupDialog.bind(key, lambda event: backupDialog.destroy())
    else:
        backupDialog = tk.Toplevel(root)
        backupDialog.wm_title("ASC Database Backup Dialog")
        #application icon
        ph1 = tk.PhotoImage(file = icon_path)
        backupDialog.iconphoto(False, ph1)
        tk.Label(backupDialog,text="Backup Procedure Successful!",font="Helvetica 12 bold").grid(row = 0, column = 0)
        tk.Label(backupDialog,text="File: " + direc + "/" + filename,font="Helvetica 12 bold").grid(row = 1, column = 0)
        okButton = tk.Button(backupDialog, text = "OK", command=backupDialog.destroy)
        okButton.grid(row = 3, column = 1, padx = 5, pady = 5)
        for key in acc_keys:
            backupDialog.bind(key, lambda event: backupDialog.destroy())



def restore():
    #restores db from a previous backup
    
    #get file for restore
    file = fd.askopenfilename(title= "Please select backup file:",
                                    filetypes= (('SQL files', '*.sql'),))
    if file == '' or file == ():
        return 1
    
    def replaceAction():
        curs = cnx.cursor()
        curs.execute("DROP TABLE IF EXISTS Appeal_Records")
        cnx.commit()
        cmd = "mysql -u root -p" + pswd + " Appeals_DB < " + file
        os.popen(cmd)
        confirmDialog.destroy()

    confirmDialog = tk.Toplevel(root)
    confirmDialog.grab_set()
    #application icon
    ph1 = tk.PhotoImage(file = icon_path)
    confirmDialog.iconphoto(False, ph1)
    tk.Label(confirmDialog,text="WARNING: The existing DB will be replaced.",font="Helvetica 12 bold",anchor="w").grid(row = 0, column = 0)
    tk.Label(confirmDialog,text="Are you sure you want to proceed?",anchor="w").grid(row = 2, column = 0)
    proceedButton = tk.Button(confirmDialog, text = "Proceed",command = replaceAction)
    proceedButton.grid(row = 4, column = 0, padx = 5, pady = 5)
    cancelButton = tk.Button(confirmDialog, text = "Cancel",command = confirmDialog.destroy)
    cancelButton.grid(row = 4, column = 2, padx = 5, pady = 5)
    for key in acc_keys:
        confirmDialog.bind(key, lambda event: confirmDialog.destroy())
    
    

#exit db
def exitdb():
    cnx.close()
    root.destroy() 

# set up main control window
root = tk.Tk(className="asc db")
root.geometry("400x400")
root.wm_title("ASC Database Control")
#application icon
ph1 = tk.PhotoImage(file = icon_path)
root.iconphoto(False, ph1)

# reports group
labelframer = tk.LabelFrame(root, text="Reports")
labelframer.pack(fill="both", expand="yes")

daterangeButton = tk.Button(labelframer, text = "Date Range Report", command = daterange)
daterangeButton.config(width = buttonsz)
daterangeButton.grid(row = 0, column = 0, padx = 5, pady = 5)

winButton = tk.Button(labelframer, text = "Specific WINs Report", command = winreport)
winButton.config(width = buttonsz)
winButton.grid(row = 0, column = 3, padx = 5, pady = 5)

winButton = tk.Button(labelframer, text = "View by WIN", command = viewWIN)
winButton.config(width = buttonsz)
winButton.grid(row = 3, column = 0, padx = 5, pady = 5)

dbstatsButton = tk.Button(labelframer, text = "DB Statistics", command = dbstats)
dbstatsButton.config(width = buttonsz)
dbstatsButton.grid(row = 3, column = 3, padx = 5, pady = 5)

# management group
labelframem = tk.LabelFrame(root, text="DB Management")
labelframem.pack(fill="both", expand="yes")

add1Button = tk.Button(labelframem, text = "Add Single Entry", command = add1)
add1Button.config(width = buttonsz)
add1Button.grid(row = 0, column = 0, padx = 5, pady = 5)

addflButton = tk.Button(labelframem, text = "Add Entries from File", command = addfl)
addflButton.config(width = buttonsz)
addflButton.grid(row = 6, column = 0, padx = 5, pady = 5)

modButton = tk.Button(labelframem, text = "Modify Existing Entry", command = modify)
modButton.config(width = buttonsz)
modButton.grid(row = 3, column = 0, padx = 5, pady = 5)

delwinButton = tk.Button(labelframem, text = "Delete Entries by WIN", command = deletewin)
delwinButton.config(width = buttonsz)
delwinButton.grid(row = 0, column = 3, padx = 5, pady = 5)

deldateButton = tk.Button(labelframem, text = "Delete Entries by Date", command = deletedate)
deldateButton.config(width = buttonsz)
deldateButton.grid(row = 3, column = 3, padx = 5, pady = 5)

# backup and restore group
labelframebu = tk.LabelFrame(root, text="Backup and Restore")
labelframebu.pack(fill="both", expand="yes")

backupButton = tk.Button(labelframebu, text = "Backup Current DB", command = backup)
backupButton.config(width = buttonsz)
backupButton.grid(row = 0, column = 0, padx = 5, pady = 5)

restoreButton = tk.Button(labelframebu, text = "Restore DB from File", command = restore)
restoreButton.config(width = buttonsz)
restoreButton.grid(row = 0, column = 3, padx = 5, pady = 5)

#exit
labelframeex = tk.LabelFrame(root, text="Exit")
labelframeex.pack(fill="both", expand="yes")
exitButton = tk.Button(labelframeex, text = "Exit DB", command = exitdb) 
exitButton.config(width = buttonsz)
exitButton.grid(row = 0, column = 3, padx = 5, pady = 5)

# main loop
root.mainloop()
