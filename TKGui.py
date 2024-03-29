from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkcalendar import DateEntry
from NarrativeGenerator import buildNarrative
from datetime import datetime
def popupText(msg:str):
    popup = Tk()
    popup.title("")
    label = ttk.Label(popup,text=msg)
    label.pack(side=TOP,fill=BOTH,padx=10,pady=10)
    ttk.Button(popup,text="Okay",command= popup.destroy).pack()
    popup.mainloop()




# Xlsx File Select
def fileBrowserXLSX(entryField):
    fileName = filedialog.askopenfilename(initialdir="/",
                                          title = "Select a File",
                                          filetypes= (("XLSX Files","*.xlsx*"),("all files","*.*")))
    entryField.delete(0,"end")
    entryField.insert(0,fileName)

# Docx File Select
def fileBrowserDOCX(entryField):
    fileName = filedialog.askopenfilename(initialdir="/",
                                          title = "Select a File",
                                          filetypes= (("DOCX Files","*.docx*"),("all files","*.*")))
    entryField.delete(0,"end")
    entryField.insert(0,fileName)

# Build Narrative
def bNar():
    saveAs = filedialog.asksaveasfile(mode='w', initialfile='New Comparison.docx',defaultextension='.docx',filetypes=[('.docx','.docx')])
    buildNarrative(old_FileName = fileA.get(),
                   new_FileName = fileB.get(),
                   dataDate = datetime.strptime(currDataDate.get(),'%m/%d/%y'),
                   previousDataDate = datetime.strptime(compDataDate.get(),'%m/%d/%y'),
                   saveLocation = saveAs.name,
                   templateLocation = templLocation.get())
    popupText("Narrative successfully created")

# Create main window
root = Tk()
root.geometry('500x350')
root.title("CPM Narrative Generator")
frm = ttk.Frame(root)
frm.pack(padx=10,pady=10,fill=BOTH,expand=TRUE)

# Declare variables
fileA = StringVar(frm,"")
fileB = StringVar(frm,"")
compDataDate = StringVar(frm,"")
currDataDate = StringVar(frm,"")
templLocation = StringVar(frm,"")

#File Select Window
ttk.Label(frm,text="Select schedules for comparison").pack(side= TOP,anchor=W,padx=5,pady=5)
fSelect = ttk.Frame(frm,borderwidth=5,relief=RAISED)
fSelect.pack(padx=5,pady=5,fill='x')

# Comprison Schedule entry
comp = ttk.Frame(fSelect)
comp.pack(fill= 'x')
ttk.Label(comp, text="Comparison Schedule").grid(row=0,column=0,sticky=W,columnspan=2)
ttk.Label(comp, text='File Path:').grid(row=1,column=0,sticky=E)
inputBoxA = ttk.Entry(comp,textvariable=fileA,width=50)
inputBoxA.grid(row=1,column=1,sticky=W)
ttk.Button(comp,text="Browse Files",command = lambda: fileBrowserXLSX(inputBoxA)).grid(row=1,column=2,padx=5)
ttk.Label(comp,text='Data Date:').grid(row=2,column=0,sticky=E)
calA = DateEntry(comp, selectmode= 'day',textvariable=compDataDate)
calA.grid(row=2,column=1,sticky=W)

# Current Schedule Entry
curr = ttk.Frame(fSelect)
curr.pack(fill= 'x',pady=5)
ttk.Label(curr, text="Current Schedule").grid(row=0,column=0,sticky=W,columnspan=2)
ttk.Label(curr, text='File Path:').grid(row=1,column=0,sticky=E)
inputBoxB = ttk.Entry(curr,textvariable=fileB,width=50)
inputBoxB.grid(row=1,column=1,sticky=W)
ttk.Button(curr,text="Browse Files",command = lambda: fileBrowserXLSX(inputBoxB)).grid(row=1,column=2,padx=5)
ttk.Label(curr,text='Data Date:').grid(row=2,column=0,sticky=E)
calB = DateEntry(curr, selectmode= 'day',textvariable= currDataDate)
calB.grid(row=2,column=1,sticky=W)

# Template File Select
templ = ttk.Frame(fSelect)
templ.pack(fill='x',pady=5)
ttk.Label(templ, text='Word Template').grid(row= 0, column= 0,sticky=W,columnspan=2)
ttk.Label(templ, text= 'File Path:').grid(row=1,column=0,sticky=E)
inputBoxT = ttk.Entry(templ,textvariable= templLocation,width= 50)
inputBoxT.grid(row=1,column=1,sticky=W)
ttk.Button(templ,text= 'Browse Files',command= lambda: fileBrowserDOCX(inputBoxT)).grid(row= 1, column=2,padx= 5)

# Build Narrative
bn = ttk.Button(frm,text="Build Narrative",command=bNar)
bn.pack(fill='x',padx=5,pady=5)

# Persistant Quit Program
ttk.Button(frm,text="Quit",command=root.destroy).pack(fill= NONE,pady=5, ipadx=5,ipady=5,anchor=E,side=BOTTOM)
root.mainloop()