import tkinter as tk
import pyodbc
import webbrowser


class GUI(tk.Frame):
    def __init__(self, master=None, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)

        w = tk.Label(self, text='Tlat Ofun Whatsapp System',font=("Helvetica", 16), justify="center")
        w.grid(row = 0, column=1)


        self.var = tk.StringVar()
        entry = tk.Entry(self, textvariable=self.var)
        entry.grid(row = 1, column=1,sticky="NE")

        btn = tk.Button(self, text='Search', command=self.read_entry)
        btn.grid(row = 1, column=1,sticky="N")


        self.mylist = []
        self.var2 = tk.StringVar(value=self.mylist)
        self.box = tk.Listbox(self, listvariable=self.var2 ,selectmode="SINGLE")
        self.box.grid(row = 1, column=1,sticky="W")
        self.box.bind('<<ListboxSelect>>', self.onselect)


        msg1 = tk.Button(self, text='שלום, אופנייך מוכנים', command=self.btn1,height = 2, width = 45)
        msg1.grid(row=4, column=1)
        msg2 = tk.Button(self, text='שלום ניסינו ליצור איתך קשר', command=self.btn2,height = 2, width = 45)
        msg2.grid(row=5, column=1)
        msg3 = tk.Button(self, text='אחר', command=self.btn3,height = 2, width = 45)
        msg3.grid(row=6, column=1)


    def btn1(self):

        phone_num = self.var.get()

        webbrowser.open('https://web.whatsapp.com/send?phone=972' + phone_num + '&text=' + "שלום, אופנייך מוכנים.נשמח אם תאסוף אותם. תלת אופן-חנות האופניים של זכרון")

    def btn2(self):


        phone_num = self.var.get()

        webbrowser.open('https://web.whatsapp.com/send?phone=972' + phone_num + '&text=' + "שלום ניסינו ליצור איתך קשר. אנא חזור אלינו. תלת אופן חנות האופניים של זכרון")

    def btn3(self):

        phone_num = self.var.get()

        webbrowser.open('https://web.whatsapp.com/send?phone=972' + phone_num)


    def read_entry(self):
        self.mylist.clear()
        x2 = self.var.get()
        temp = 'select name from main where name like '
        temp3 = ' or phone like'
        temp2 = temp + "'%" + x2 + "%'" + temp3 + "'%" + x2 + "%'"

        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\db\db.accdb;')
        cursor = conn.cursor()
        cursor.execute(temp2)

        index = 0

        for row in cursor.fetchall():
            self.mylist.append(row[0])
            index = index + 1

        self.var2.set(self.mylist)

    def onselect(self,event):

        # Note here that Tkinter passes an event object to onselect()
        w = event.widget
        index = int(w.curselection()[0])
        value = w.get(index)

        temp = 'select phone from main where name ='
        sql = temp + "'" + value + "'"
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\db\db.accdb;')
        cursor = conn.cursor()
        cursor.execute(sql)

        for row in cursor.fetchall():
            p = row[0]

        self.var.set(p)


def main():
    root = tk.Tk()
    root.geometry('330x330')
    win = GUI(root)
    win.grid()
    root.title("Tlat Ofun")
    root.mainloop()

if __name__ == '__main__':
    main()