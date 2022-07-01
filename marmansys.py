from tkinter import *
import xlrd
import xlwt
from xlutils.copy import copy
from tkinter import messagebox
from tkinter.simpledialog import askfloat
from tkinter.simpledialog import askstring
from tkinter.simpledialog import askinteger
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from threading import Thread
matplotlib.use('Qt5Agg')
from re import fullmatch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pretty_html_table import build_table



def load():
        global book
        book=xlrd.open_workbook("/media/bkm/F014230D1422D67E/MACHINE LEARNING-TECHVANTO/MARKS-MANAGEMENT-SYSTEM/mms.xls")
        global sheet
        sheet=book.sheet_by_index(0)
        global wb
        wb=copy(book)
        global wsheet
        wsheet=wb.get_sheet(0)
        global mails
        mails = [sheet.cell_value(k,1) for k in range(1,sheet.nrows)]
def save():
    wb.save('/media/bkm/F014230D1422D67E/MACHINE LEARNING-TECHVANTO/MARKS-MANAGEMENT-SYSTEM/mms.xls')
def startsession():
    global s
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login('bharatkumarmareedu@gmail.com', open('/media/bkm/F014230D1422D67E/MACHINE LEARNING-TECHVANTO/MARKS-MANAGEMENT-SYSTEM/secret.txt','r').read())
    print("Logged In")
class main(Tk):
    def __init__(self):
        super().__init__()
        self.title("Marks Management System")
        self.geometry("700x700")
        self.resizable(0,0)
        Button(self,text="Student Wise Results",font=("TimesNewRoman"),command=lambda:self.stuwres(self),height=5,width=20).place(x=100,y=40)
        Button(self,text="Subject Wise Results",font=("TimesNewRoman"),command=lambda:self.subwres(self),height=5,width=20).place(x=400,y=40)
        Button(self,text="Subject Wise Comparision",font=("TimesNewRoman"),command=lambda:self.subwcom(self),height=5,width=20).place(x=100,y=220)
        Button(self,text="Send all reports to Parents",font=("TimesNewRoman"),command=self.sendreports,height=5,width=20).place(x=400,y=220)
        Button(self,text="Send weak student reports",font=("TimesNewRoman"),height=5,width=20,command=self.sweaksreports).place(x=100,y=400)
        Button(self,text="Add or Update Marks",font=("TimesNewRoman"),height=5,width=20,command=lambda:self.addmarks(self)).place(x=400,y=400)
        Button(self,text="Send Report to Teachers",font=("TimesNewRoman"),height=5,width=20,command=self.sreptchrs).place(x=100,y=580)
        Button(self,text="EXIT",height=5,font=("TimesNewRoman"),width=20,bg="red",activebackground="#f24438",fg="white",activeforeground="White",command=lambda:self.destroy()).place(x=400,y=580)
    def sreptchrs(self):
        startsession()
        load()
        def runthread():
            sheet2 = book.sheet_by_index(1)
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            marks = []
            emails = [sheet2.cell_value(k,1) for k in range(1,sheet2.nrows)]
            for i in range(len(emails)):
                marks = [sheet.cell_value(j,i+2) for j in range(1,sheet.nrows)]
                df = pd.DataFrame(data={"Student Name":names,"Marks":marks})
                html = '''
                        <html>
                            <body>
                                <h1>Report of {0}</h1>
                                  {1}
                                    </body>
                                </html>
                                '''.format(sheet.cell_value(0,i+2),build_table(df, 'blue_light'))
                email_message = MIMEMultipart()
                email_message['From'] = "bharatkumarmareedu@gmail.com"
                email_message['To'] = emails[i]
                email_message['Subject'] = "Report Of Your Subject"
                email_message.attach(MIMEText(html, "html"))
                email_string = email_message.as_string()
                s.sendmail("bharatkumarmareedu@gmail.com",emails[i],email_string)
        t1=Thread(target=runthread)
        t1.start()
        t1.join()
        messagebox.showinfo("Success","Report Sent To Teachers")
        s.quit()
    def sweaksreports(self):
        startsession()
        load()
        def runthread():
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            inde = ['Maths','Physics','Chemistry','Biology','English','Hindi','Total','Percentage']
            emails = [sheet.cell_value(k,1) for k in range(1,sheet.nrows)]
            for i in range(len(names)):
                if(int(sheet.cell_value(i+1,9))<60):
                    marks = [sheet.cell_value(i+1,j+2) for j in range(8)]

                    df = pd.DataFrame(data={"Subject":inde,"Marks":marks})
                    html = '''
                                    <html>
                                        <body>
                                            <h1>Report of {0}</h1>
                                            <h2>Your ward is a weak student</h2>
                                            {1}
                                        </body>
                                    </html>
                                    '''.format(names[i],build_table(df, 'blue_light'))
                    email_message = MIMEMultipart()
                    email_message['From'] = "bharatkumarmareedu@gmail.com"
                    email_message['To'] = emails[i]
                    email_message['Subject'] = "Report Card"
                    email_message.attach(MIMEText(html, "html"))
                    email_string = email_message.as_string()
                    s.sendmail("bharatkumarmareedu@gmail.com",emails[i],email_string)
        t1=Thread(target=runthread)
        t1.start()
        t1.join()
        messagebox.showinfo("Success","Report Cards of Weak Students Sent Successfully")
        s.quit()
    def subwcom(self,s):
        x1 = [2,4,6,8,10,12]
        y1 = []
        load()
        for j in range(6):
            y1.append(np.mean(np.array([sheet.cell_value(i+1,j+2) for i in range(sheet.nrows-1)])))
        plt.bar(x1,y1,tick_label=['Maths','Physics','Chemistry','Biology','English','Hindi'],color='green',width=0.8)
        plt.plot(x1,[35 for i in range(6)],linestyle='dashed',label='Fail',color='red')
        plt.plot(x1,[60 for i in range(6)],linestyle='dashed',color='orange',label='weak')
        plt.ylim(0,100)
        plt.xlabel('Subjects')
        plt.ylabel('Mean Marks')
        plt.title('Subject Wise Comparison')
        for j in range(6):
            plt.text(x1[j],y1[j],'%.1f'%y1[j])
        plt.legend()
        plt.show()
    def sendreports(self):
        startsession()
        load()
        def runthread():
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            inde = ['Maths','Physics','Chemistry','Biology','English','Hindi','Total','Percentage']
            emails = [sheet.cell_value(k,1) for k in range(1,sheet.nrows)]
            for i in range(len(names)):
                marks = [sheet.cell_value(i+1,j+2) for j in range(8)]

                df = pd.DataFrame(data={"Subject":inde,"Marks":marks})
                html = '''
                                <html>
                                    <body>
                                        <h1>Report of {0}</h1>
                                        {1}
                                    </body>
                                </html>
                                '''.format(names[i],build_table(df, 'blue_light'))
                email_message = MIMEMultipart()
                email_message['From'] = "bharatkumarmareedu@gmail.com"
                email_message['To'] = emails[i]
                email_message['Subject'] = "Report Card"
                email_message.attach(MIMEText(html, "html"))
                email_string = email_message.as_string()
                s.sendmail("bharatkumarmareedu@gmail.com",emails[i],email_string)
        t1=Thread(target=runthread)
        t1.start()
        t1.join()
        messagebox.showinfo("Success","Report Cards Sent Successfully")
        s.quit()
    def subwres(self,s):
        s.withdraw()
        f=Toplevel(s)
        f.grab_set()
        f.geometry("300x400")
        f.resizable(0,0)
        f.protocol("WM_DELETE_WINDOW",lambda:(s.deiconify(), f.destroy()))
        def subwmar():
            if(sub.get()==0):
                return
            load()
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            x1=[k*2 for k in range(1,len(names)+1)]
            marks = [sheet.cell_value(k,int(sub.get())) for k in range(1,sheet.nrows)]
            plt.bar(x1,marks,tick_label=names,color='green',width=0.8)
            plt.plot(x1,[35 for i in range(len(names))],linestyle='dashed',label='Fail',color='red')
            plt.plot(x1,[60 for i in range(len(names))],linestyle='dashed',color='orange',label='weak')
            plt.ylim(0,100)
            plt.xlabel('Students')
            plt.ylabel('Marks')
            plt.title('Results of %s'%sheet.cell_value(0,int(sub.get())))
            for j in range(len(names)):
                plt.text(x1[j],marks[j],'%.1f'%marks[j])
            plt.legend()
            plt.show()
        Label(f,text="Slect Subject To Display",font=("TimesNewRoman")).place(anchor=CENTER,x=150,y=20)
        sub=IntVar()
        Radiobutton(f,text="Maths",variable=sub,value=2).place(x=120,y=50)
        Radiobutton(f,text="Physics",variable=sub,value=3).place(x=120,y=80)
        Radiobutton(f,text="Chemistry",variable=sub,value=4).place(x=120,y=110)
        Radiobutton(f,text="Biology",variable=sub,value=5).place(x=120,y=140)
        Radiobutton(f,text="English",variable=sub,value=6).place(x=120,y=170)
        Radiobutton(f,text="Hindi",variable=sub,value=7).place(x=120,y=200)
        Button(f,text="Submit",font=("TimesNewRoman"),bg="Green",fg="White",activebackground="Green",command=subwmar).place(x=70,y=250)
        Button(f,text="Cancel",font=("TimesNewRoman"),bg="Red",fg="White",activebackground="Red",command=lambda:(f.destroy(),s.deiconify())).place(x=170,y=250)
        
    def stuwres(self,s):
        s.withdraw()
        email = askstring("E-Mail", "Enter Parent's Mail : ")
        if(email!=None):
            load()
            if(email in mails):
                r = mails.index(email)
                x1=[2,4,6,8,10,12,14]
                y1=[sheet.cell_value(r+1,i) for i in range(2,8)]
                y1.append(sheet.cell_value(r+1,9))
                plt.bar(x1,y1,tick_label=['Maths','Physics','Chemistry','Biology','English','Hindi','Percentage'],color='green',width=0.8)
                plt.plot(x1,[35 for i in range(7)],linestyle='dashed',label='Fail',color='red')
                plt.plot(x1,[60 for i in range(7)],linestyle='dashed',color='orange',label='weak')
                plt.ylim(0,100)
                plt.xlabel('Subjects')
                plt.ylabel('Marks')
                plt.title('Results of %s'%sheet.cell_value(r+1,0))
                for j in range(7):
                    plt.text(x1[j],y1[j],'%.1f'%y1[j])
                plt.legend()
                plt.show()
            else:
                messagebox.showinfo("Invalid", "No Such Mail Exist !!")
        s.deiconify()
    def addmarks(self,s):
        s.withdraw()
        f=Toplevel(s)
        def update():
            load()
            m =  askstring("E-Mail", "Enter Parent's Mail : ")
            if(m!=None):
                if(m in mails):
                    up = Toplevel(f)
                    def updsub():
                        newmark = askfloat("Update Marks", "Enter New Marks ",parent=up)
                        if(newmark!=None):
                            if(newmark>100 or newmark<0):
                                messagebox.showwarning("Invalid","Marks range is 0-100")
                                return
                            wsheet.write(r,int(sub.get()),float(newmark))
                            t = np.sum(np.array([float(sheet.cell_value(r,i)) for i in range(2,8)]))
                            wsheet.write(r,8,t)
                            wsheet.write(r,9,t/6)
                            save()
                            messagebox.showinfo("Saved","Marks Updated Successfully")
                    up.grab_set()
                    up.geometry("300x400")
                    up.resizable(0,0)
                    up.protocol("WM_DELETE_WINDOW",lambda:(s.deiconify(), f.destroy()))
                    r = mails.index(m)+1
                    Label(up,text="Slect Subject To Update",font=("TimesNewRoman")).place(anchor=CENTER,x=150,y=20)
                    sub=IntVar()
                    Radiobutton(up,text="Maths",variable=sub,value=2).place(x=120,y=50)
                    Radiobutton(up,text="Physics",variable=sub,value=3).place(x=120,y=80)
                    Radiobutton(up,text="Chemistry",variable=sub,value=4).place(x=120,y=110)
                    Radiobutton(up,text="Biology",variable=sub,value=5).place(x=120,y=140)
                    Radiobutton(up,text="English",variable=sub,value=6).place(x=120,y=170)
                    Radiobutton(up,text="Hindi",variable=sub,value=7).place(x=120,y=200)
                    Button(up,text="Submit",font=("TimesNewRoman"),bg="Green",fg="White",activebackground="Green",command=updsub).place(x=70,y=250)
                    Button(up,text="Cancel",font=("TimesNewRoman"),bg="Red",fg="White",activebackground="Red",command=lambda:up.destroy()).place(x=170,y=250)
                else:
                    messagebox.showinfo("Invalid","No Such Mail Exist !!")
            
        def cupdate():
            try:
                vals = np.array([float(m.get()),float(p.get()),float(c.get()),float(b.get()),float(e.get()),float(h.get())])
            except ValueError:
                messagebox.showwarning("Mandatory", "All Feilds are Mandatory !\nMarks should be integers",parent=f)
                return
            if(name.get()=='' or email.get()==''):
                messagebox.showwarning("Mandatory","Email and Name Feilds are Mandatory",parent=f)
                return
            if(np.any(vals>100) or np.any(vals<0)):
                messagebox.showwarning("Invalid","Marks range is 0-100",parent=f)
                return
            if(not fullmatch(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email.get())):
                messagebox.showinfo("Invalid mail", "Enter a Valid Mail")
                return
            if(not messagebox.askyesno("Confirmation","Are you sure to Save ? ")):
                return
            load()
            r = sheet.nrows
            wsheet.write(r,0,name.get())
            wsheet.write(r,1,email.get())
            wsheet.write(r,2,float(m.get()))
            wsheet.write(r,3,float(p.get()))
            wsheet.write(r,4,float(c.get()))
            wsheet.write(r,5,float(b.get()))
            wsheet.write(r,6,float(e.get()))
            wsheet.write(r,7,float(h.get()))
            t = np.sum(np.array([float(m.get()),float(p.get()),float(c.get()),float(b.get()),float(e.get()),float(h.get())]))
            wsheet.write(r,8,t)
            wsheet.write(r,9,t/6)
            save()
            messagebox.showinfo("Success","Marks Added Successfully!")
            name.delete(0,END)
            p.delete(0,END)
            m.delete(0,END)
            c.delete(0,END)
            b.delete(0,END)
            e.delete(0,END)
            h.delete(0,END)
            email.delete(0,END)
        f.grab_set()
        f.geometry("400x500")
        f.resizable(0,0)
        f.protocol("WM_DELETE_WINDOW",lambda:(s.deiconify(), f.destroy()))

        Label(f,text="Enter Your Name : ",font=("TimesNewRoman")).place(x=20,y=30)
        name=Entry(f,width=30)
        name.place(x=150,y=30)
        Label(f,text="Parent Email : ",font=("TimesNewRoman")).place(x=20,y=60)
        email=Entry(f,width=30)
        email.place(x=150,y=60)
        Label(f,text="Maths Marks : ",font=("TimesNewRoman")).place(x=20,y=150)
        m=Entry(f,width=10)
        m.place(x=150,y=150)
        Label(f,text="Physics Marks : ",font=("TimesNewRoman")).place(x=20,y=180)
        p=Entry(f,width=10)
        p.place(x=150,y=180)
        Label(f,text="Chemistry Marks : ",font=("TimesNewRoman")).place(x=20,y=210)
        c=Entry(f,width=10)
        c.place(x=150,y=210)
        Label(f,text="Biology Marks : ",font=("TimesNewRoman")).place(x=20,y=240)
        b=Entry(f,width=10)
        b.place(x=150,y=240)
        Label(f,text="English Marks : ",font=("TimesNewRoman")).place(x=20,y=270)
        e=Entry(f,width=10)
        e.place(x=150,y=270)
        Label(f,text="Hindi Marks : ",font=("TimesNewRoman")).place(x=20,y=300)
        h=Entry(f,width=10)
        h.place(x=150,y=300)
        Button(f,text="Submit",font=("TimesNewRoman"),bg="Green",fg="White",activebackground="Green",command=cupdate).place(x=70,y=400)
        Button(f,text="Cancel",font=("TimesNewRoman"),bg="Red",fg="White",activebackground="Red",command=lambda:(f.destroy(),s.deiconify())).place(x=250,y=400)
        Button(f,text="Update Marks",font=("TimesNewRoman"),bg="Yellow",activebackground="Yellow",command=update).place(x=135,y=350)
main().mainloop()