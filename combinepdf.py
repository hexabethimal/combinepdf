import os
import sys
from pypdf import PdfWriter
from docx2pdf import convert
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

folderWithPDFs = ""
destinationFolder = ""
newPDFname = ""
         
def select_source_folder():
    x = fd.askdirectory()  
    entry1.delete(0, tk.END)
    entry1.insert(0, x)     
    
    if entry1.get() != "":
        entry1.configure(state="readonly")

def select_destination_folder():
    x = fd.askdirectory()   
    entry2.delete(0, tk.END)
    entry2.insert(0, x)  
    
    if entry2.get() != "":
        entry2.configure(state="readonly")
    
def startCheck(): 
    folderWithPDFs = entry1.get()
    destinationFolder = entry2.get()
    newPDFname = entry3.get()
    
    if folderWithPDFs and destinationFolder and newPDFname:
        mergePDFs(folderWithPDFs, destinationFolder, newPDFname)
        showinfo("Finished", "New PDF Created")
        sys.exit()  
    else:
        showinfo("Invalid Selection","Select a source folder, destination folder, and choose a new PDF filename first.")

def mergePDFs(folderWithPDFs, destinationFolder, newPDFname):        
    filesList = os.listdir(folderWithPDFs)
    fileDir = folderWithPDFs + "/"            
    newList = [fileDir + x for x in filesList]
    merger = PdfWriter()
    createdFiles = []
    
    if var1.get() == 1:
        for pdf in newList:          
            if ".docx" in pdf:
                    convert(pdf)  
                    createdFiles.append(pdf.replace(".docx", ".pdf")) 
        filesList = os.listdir(folderWithPDFs)
        newList = [fileDir + x for x in filesList]
    
    try:
        for pdf in newList: 
            if ".pdf" in pdf:            
                merger.append(pdf)
            
        if ".pdf" in newPDFname:
            newPDFname = newPDFname.replace(".pdf", "")
            
        merger.write(destinationFolder + "/" + newPDFname + ".pdf")
        merger.close()
    except:
        showinfo("Unable to combine. 1 or more PDFs may be password protected.")

    if var1.get() == 1:
        for x in range(len(createdFiles)):                            
            try:
                os.remove(createdFiles[x])
            except OSError:
                pass

window = tk.Tk()
window.title("Combine PDFs")
window.geometry("450x390")
window.resizable(False, False)
window.configure(bg="#AD9A9D")
style = ttk.Style(window)
style.configure("TButton", font=('wasy10', 14))
var1 = tk.IntVar()
getSourceFolder_button = ttk.Button(window, text="Select Source Folder", command=select_source_folder, style="TButton")
getSourceFolder_button.pack(pady=(15,0))     
entry1 = tk.Entry(font=('wasy10', 12))
entry1.pack(ipadx=90, ipady=5, pady=15)
getDestFolder_button = ttk.Button(window, text="Select Destination Folder", command=select_destination_folder, style="TButton")
getDestFolder_button.pack()     
entry2 = tk.Entry(font=('wasy10', 12))
entry2.pack(ipadx=90, ipady=5, pady=15)
label1 = tk.Label(text="What do you want your new PDF file name to be?", font=('wasy10', 14))
label1.config(bg="#AD9A9D")
label1.pack()
entry3 = tk.Entry(justify="center", font=('wasy10', 12))
entry3.pack(ipadx=0, ipady=5, pady=15)
checkbox1 = tk.Checkbutton(window, text="Convert & include '.docx' files in final PDF?", font=('wasy10', 14), variable=var1, onvalue=1, offvalue=0)
checkbox1.config(bg="#AD9A9D")
checkbox1.pack(pady=(0,15))
mergePDF_button = ttk.Button(window, text="Start", command=startCheck)
mergePDF_button.pack()
window.mainloop()
