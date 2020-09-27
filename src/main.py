from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename
import os
import PyPDF2


def open_file():
    global file
    file = askopenfilename(defaultextension=".pdf", filetypes=[("All Files", "*.*"), ("PDF files", "*.pdf")])
    if file == "":
        file = None
    else:
        root.title(os.path.basename(file) + "-PDF to Word")
        t1.delete(1.0, END)
        e1.delete(0, END)
        e1.insert(0, file)


def convert():
    global file
    f = open(file, "rb")
    pdf_reader = PyPDF2.PdfFileReader(f)
    pages = pdf_reader.numPages
    x = 0
    for x in range(x, pages):
        page = pdf_reader.getPage(x)
        text = page.extractText()
        t1.insert(INSERT, text)
    f.close()


def save_file():
    global file
    file = asksaveasfilename(initialfile="Untitled.word", defaultextension=".word",
                             filetypes=[("Word Documents", "*.word"), ("Text Documents", "*.text")])
    if file == "":
        file = None
    else:
        f = open(file, "w")
        f.write(t1.get(1.0, END))
        f.close()


root = Tk()
root.geometry("600x470")
root.wm_iconbitmap("F:\Tkinter\pdf to word converter\src\icon.jpg")
root.title("PDF To Word- convert your PDF to Word")
l1 = Label(root, text="Select Pdf file", font=" Lucinda 15 ")
l1.place(x=70, y=50)
e1 = Entry(root, text=" ", font=" Lucinda 15 ")
e1.place(x=200, y=50)
b1 = Button(root, text="Open", font="Lucinda 11", command=open_file)
b1.place(x=440, y=50)
b2 = Button(root, text="   Convert  ", font="  lucinda 15 ", command=convert)
b2.place(x=250, y=85)
file = None
t1 = Text(root, font=" lucinda 10", relief="groove")
t1.place(x=0, y=130, width=600, height=300)
sb = Scrollbar(t1, cursor="arrow")
sb.pack(side=RIGHT, fill=Y)
sb.config(command=t1.yview)
t1.config(yscrollcommand=sb.set)
b3 = Button(root, text="Save", font="lucinda 15", command=save_file)
b3.place(x=0, y=430)
root.mainloop()
