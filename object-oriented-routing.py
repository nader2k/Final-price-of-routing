from tkinter import *
import matplotlib.pyplot as plt
from tkinter import messagebox
import tkinter as tk
import tkinter.scrolledtext as st
import os
import sys
import pandas as pd
i = 0
flag = 0
flag1 = 0
flag2 = 0
flag3 = 0
flag4 = 0
flag5 = 0
flag8 = 0
n = 0
select_method = 0
final_price = 0
workpiece_name = []

class SetScreen:

    def __init__(self, root_name, y_offset, flag_n=0):
        self.root_name = root_name
        self.y_offset = y_offset
        self.flag = flag_n
        self.root_name.update_idletasks()
        self.w_a = self.root_name.winfo_reqwidth()
        self.h_a = self.root_name.winfo_reqheight()
        self.ws_a = self.root_name.winfo_screenwidth()
        self.hs_a = self.root_name.winfo_screenheight()
        self.x = (self.ws_a / 2) - (self.w_a / 2)
        if self.flag:
            self.y = (self.hs_a - self.h_a) / 5
        else:
            self.y = (self.hs_a / 2) - (self.h_a / 2) - (self.hs_a / self.y_offset)
        self.root_name.geometry('+%d+%d' % (self.x, self.y))

class IsNumber:

    def __init__(self, num):
        self.out = 0
        self.num = num
        self.out = None
        self.check()

    def __int__(self):
        return self.out

    def check(self):
        try:
            float(self.num)
            self.out = 1
        except ValueError:
            self.out = 0


class InputCheck:
    def __init__(self, num, text):
        self.num = num
        self.text = text
        self.x = None
        self.check()

    def __iter__(self):
        return iter(self.x)

    def check(self):
        self.num = self.num.strip()
        if self.num:
            if int(IsNumber(self.num)):
                self.num = float(self.num)
                if self.num > 0:
                    self.x = [1, self.num]
                    return self.x
                else:
                    self.x = [0, "{} باید عددی بزرگتر از صفر باشد".format(self.text)]
                    return self.x
            else:
                self.x = [0, "در قسمت {} باید عدد وارد شود".format(self.text)]
                return self.x
        else:
            self.x = [0, "{}نمی تواند خالی باشد ".format(self.text)]
            return self.x


class SelectInputMethod:
    def __init__(self):
        self.master = None
        self.v = None
        self.newframe = None
        self.creat_new_window()

    def save_select(self):
        global select_method
        select_method = self.v.get()
        select_method = int(select_method)
        if select_method == 0:
            messagebox.showinfo("Message", " یک گزینه را انتخاب کنید ")
        if select_method == 1 or select_method == 2:
            self.master.destroy()

    def creat_new_window(self):
        self.master = Tk()
        self.master.title("انتخاب روش ورود اطلاعات")
        self.master.resizable(width=False, height=False)
        self.v = IntVar()
        Label(self.master, text="انتخاب روش ورود اطلاعات", justify=CENTER, padx=20, font="arial 14 bold").grid(row=0,
                                                                                                               column=0,
                                                                                                               padx=20,
                                                                                                               pady=(10, 15))
        Radiobutton(self.master, text="ورود اطلاعات به صورت گرافیکی", justify=LEFT, padx=20, variable=self.v, value=1,
                    font="bnazanin 10").grid(row=1, column=0)
        Radiobutton(self.master, text="ورود اطلاعات به وسیله فایل اکسل", justify=LEFT, padx=20, variable=self.v,
                    value=2,
                    font="bnazanin 10").grid(row=2, column=0)
        self.newframe = Frame(self.master)
        self.newframe.grid(row=3, column=0, sticky="nsew")
     
        SetScreen(self.master, 5.5)
        self.master.mainloop()


class ExcelDataEntry:

    def __init__(self, excel_file_name='Wc_Information.xlsx'):
        self.excel_file_name = excel_file_name
        self.master = None
        self.v = None
        self.myframe = None
        self.text_area = None
        self.text = None
        self.newframe = None
        self.k1 = None
        self.newframe2 = None
        self.w = None
        self.new_window()

    def open_excel(self):
        Label(self.master, text="منتظر بمانيد، فايل اکسل باز شود", justify=LEFT, fg='red',
              font="bnazanin 14 bold").grid(row=0, column=0, padx=0, pady=0)
        os.system('start excel.exe "{}"'.format(self.excel_file_name))

    def next_level(self):
        global flag2, workpiece_name, flag4
        self.w = int(ExcelDataRead('Wc_Information.xlsx'))
        if not self.w:
            flag4 = 1
            flag2 = 2
            workpiece_name = self.k1.get()
            self.master.destroy()

    def new_window(self):
        self.master = Tk()
        self.master.title("ورود اطلاعات به وسيله اکسل")
        # master.geometry('+380+135')
        self.master.resizable(width=False, height=False)
        self.v = IntVar()
        Label(self.master, text="ورود اطلاعات به وسيله اکسل", justify=LEFT, font="bnazanin 12 bold").grid(row=0,
                                                                                                          column=0,
                                                                                                          padx=85,
                                                                                                          pady=20)

        self.myframe = Frame(self.master, relief=GROOVE, width=60, height=40, bd=1)
        self.myframe.grid(padx=3)
        self.text_area = st.ScrolledText(self.myframe, width=58, height=6, font=("b nazanin", 12))
        self.text_area.grid(row=2, column=0)

        # Inserting Text which is read only,
        text_row1 = '.الف) ابتدا بر روی دکمه باز کردن فایل اکسل ، کلیک نمایید\n'
        text_row2 = '.ب) در فایل اکسل مربوطه ، به تعداد ایستگاه های کاری مورد نیاز ، شیت ایجاد نمایید\n'
        text_row3 = 'ج) اطلاعات مربوط به هر ایستگاه کاری را در یک شیت جداگانه ( طبق فرمت مثال\n'
        text_row4 = '.موجود در فایل اکسل ) وارد نمایید\n'
        text_row5 = 'د) به این نکته توجه شود که در هر شیت، ردیف اول مربوط به نام ستون ها می باشد\n'
        text_row6 = '.و اگر در این ردیف اطلاعاتی وارد شود، توسط نرم افزار خوانده نمی شود\n'
        text_row7 = 'ه) پس از ورود اطلاعات در فایل اکسل و ذخیره نمودن آن، بر روی دکمه ادامه کلیک\n'
        text_row8 = '.نمایید'
        self.text = text_row1 + text_row2 + text_row3 + text_row4 + text_row5 + text_row6 + text_row7 + text_row8

        self.text_area.tag_configure('tag-right', justify='right')
        self.text_area.insert(END, self.text, 'tag-right')

        # Making the text read only
        self.text_area.configure(state='disabled')

        self.newframe = Frame(self.master)
        self.newframe.grid(row=3, column=0)
        Label(self.newframe, text=": نام نیمه ساخته ", font="bnazanin 12 bold", anchor='e').grid(row=0, column=1,
                                                                                                 padx=(60, 0),
                                                                                                 pady=(20, 0),
                                                                                                 sticky='e')
        self.k1 = Entry(self.newframe, width=12, justify=LEFT)
        self.k1.grid(row=0, column=0, sticky=W, pady=(20, 0))

        self.newframe2 = Frame(self.master)
        self.newframe2.grid(row=4, column=0)
        Button(self.newframe2, text="خروج", command=self.master.destroy, padx=12, pady=4, font="bnazanin 10 bold",
               background='#eaeaea').grid(row=0, column=2, padx=35, pady=(20, 25))
        Button(self.newframe2, text="باز کردن فایل اکسل", command=self.open_excel, padx=4, pady=4,
               font="bnazanin 10 bold",
               background='#eaeaea').grid(row=0, column=1, padx=35, pady=(20, 25))
   

        SetScreen(self.master, 7)
        mainloop()


class ExcelDataRead:

    def __init__(self, excel_file_name='Wc_Information.xlsx'):
        self.excel_file_name = excel_file_name
        self.xl = None
        self.sheet_name = None
        self.number_of_sheet = None
        self.sheets = None
        self.temp1 = None
        self.temp2 = None
        self.temp3 = None
        self.temp4 = None
        self.temp5 = None
        self.temp6 = None
        self.falg8 = None
        self.read()

    def __int__(self):
        return self.flag8

    def read(self):

        global flag2, all_work_center, flag8
        self.xl = pd.ExcelFile(self.excel_file_name)
        self.sheet_name = self.xl.sheet_names
        self.number_of_sheet = len(self.sheet_name)

        self.sheets = []
        with self.xl as x:
            for h in range(0, self.number_of_sheet):
                self.sheets.append(pd.read_excel(x, self.sheet_name[h]))

        all_work_center = []
        flag8 = 0
        for g in range(0, self.number_of_sheet):
            self.temp1 = []
            self.temp2 = []
            self.temp3 = []

            temp = self.sheets[g].iloc[:, 0]
            for p in range(len(temp)):
                self.temp1.append(str(temp[p]))

            temp = self.sheets[g].iloc[:, 1]
            for p in range(len(temp)):
                if temp[p] != temp[p]:
                    check = list(InputCheck(' ', 'تعداد واحد مصرفی '))
                else:
                    check = list(InputCheck(str(temp[p]), 'تعداد واحد مصرفی '))
                if check[0] == 0:
                    flag8 = 1
                    messagebox.showinfo("Message", check[1])
                    break
                else:
                    self.temp2.append(check[1])

            temp = self.sheets[g].iloc[:, 2]
            for p in range(len(temp)):
                if temp[p] != temp[p]:
                    check = list(InputCheck(' ', 'قیمت یک واحد قطعه '))
                else:
                    check = list(InputCheck(str(temp[p]), 'قیمت یک واحد قطعه '))
                if check[0] == 0:
                    flag8 = 1
                    messagebox.showinfo("Message", check[1])
                    break
                else:
                    self.temp3.append(check[1])

            temp = self.sheets[g].iloc[:, 3]
            if temp[0] != temp[0]:
                check = list(InputCheck(' ', 'زمان عملیات '))
            else:
                check = list(InputCheck(str(temp[0]), 'زمان عملیات '))
            if check[0] == 0:
                flag8 = 1
                messagebox.showinfo("Message", check[1])
            else:
                self.temp4 = round(check[1], 2)

            temp = self.sheets[g].iloc[:, 4]
            if temp[0] != temp[0]:
                check = list(InputCheck(' ', 'هزینه واحد زمان '))
            else:
                check = list(InputCheck(str(temp[0]), 'هزینه واحد زمان '))
            if check[0] == 0:
                flag8 = 1
                messagebox.showinfo("Message", check[1])
            else:
                self.temp5 = round(check[1], 2)

            temp = self.sheets[g].iloc[:, 5]
            if temp[0] != temp[0]:
                check = list(InputCheck(' ', 'درصد هزینه سربار '))
            else:
                check = list(InputCheck(str(temp[0]), 'درصد هزینه سربار '))
            if check[0] == 0:
                flag8 = 1
                messagebox.showinfo("Message", check[1])
            else:
                self.temp6 = round(check[1], 2)

            if flag8 == 1:
                self.temp1 = []
                self.temp2 = []
                self.temp3 = []
                all_work_center = []
                break
            else:
                all_work_center.append([self.temp3, self.temp2, self.temp1, self.temp5, self.temp4, self.temp6])
        self.flag8 = flag8


class WcNumber:

    def __init__(self):
        self.master = None
        self.newframe = None
        self.k1 = None
        self.newframe2 = None
        self.k2 = None
        self.newframe3 = None
        self.creat_window()

    def save_wc_num(self):
        global n, workpiece_name, flag1
        workpiece_name = self.k1.get()
        b = self.k2.get()
        check = list(InputCheck(b, 'تعداد ایستگاه ها '))
        if check[0] == 0:
            messagebox.showinfo("Message", check[1])
        else:
            if check[1] == round(check[1]):
                n = int(check[1])
                flag1 = 1
                self.master.destroy()
            else:
                messagebox.showinfo("Message", "تعداد ایستگاه ها باید عدد صحیح باشد")

    def creat_window(self):
        self.master = Tk()
        self.master.title("محاسبه قیمت نیمه ساخته")
        self.master.resizable(width=False, height=False)

        Label(self.master, text="محاسبه قیمت نهایی یک نیمه ساخته", justify=LEFT, font="bnazanin 13 bold").grid(row=0,
                                                                                                               column=0,
                                                                                                               padx=20,
                                                                                                               pady=15)
        self.newframe = Frame(self.master)
        self.newframe.grid(row=1, column=0, sticky="nsew")
        Label(self.newframe, text="نام نیمه ساخته", font="bnazanin 11").grid(row=0, column=1, padx=(107, 30), pady=15)
        self.k1 = Entry(self.newframe, width=10)
        self.k1.grid(row=0, column=0, padx=(30, 0), pady=15)

        self.newframe2 = Frame(self.master)
        self.newframe2.grid(row=2, column=0, sticky="nsew")
        Label(self.newframe2, text="تعداد ایستگاه های کاری", font="bnazanin 11").grid(row=0, column=1, padx=(50, 30),
                                                                                      pady=15)
        self.k2 = Entry(self.newframe2, width=10)
        self.k2.insert(10, 2)
        self.k2.grid(row=0, column=0, padx=(30, 0), pady=15)

        self.newframe3 = Frame(self.master)
        self.newframe3.grid(row=3, column=0, sticky="nsew")
  
        SetScreen(self.master, 5.5)
        mainloop()


class WorkCenter:

    def __init__(self, wc_number):
        global i, flag5
        flag5 = 1
        i = 0
        self.wc_number = wc_number
        self.e1 = [0] * 200
        self.e2 = [0] * 200
        self.e3 = [0] * 200
        self.unit_price = []
        self.consumption = []
        self.name = []
        self.m_width = 485 + 16 + 100
        self.m_height = 300 + 40 + 40

        self.flag7 = None
        self.t1 = None
        self.t2 = None
        self.t3 = None
        self.m1 = None
        self.root = None
        self.var1 = None
        self.myframe = None
        self.canvas = None
        self.frame = None
        self.myscrollbar = None
        self.h1 = None
        self.h2 = None
        self.h3 = None

        self.creat_window()

    def __iter__(self):
        return iter([self.unit_price, self.consumption, self.name, self.t1, self.t2, self.t3])

    def creat_e(self):
        global i
        self.e1[i] = Entry(self.frame, width=10)
        self.e2[i] = Entry(self.frame, width=5)
        self.e3[i] = Entry(self.frame, width=15)
        # self.e[i].insert(10, "")
        # self.e2[i].insert(10, "")
        # self.e3[i].insert(10,"")
        self.e1[i].grid(row=5 + i, column=0, pady=8, padx=40)
        self.e2[i].grid(row=5 + i, column=1, pady=8, padx=(37, 4))
        self.e3[i].grid(row=5 + i, column=2, pady=8, padx=56)

    def save_fields(self):
        global i, flag5
        self.t1 = 1
        self.t2 = 1
        self.t3 = 0
        # nonlocal unit_price , consumption , name, t1, t2, t3
        self.flag7 = 0
        for d in range(0, i + 1):
            temp1 = self.e1[d].get()
            temp2 = self.e2[d].get()
            temp3 = self.e3[d].get()

            self.name.append(temp3)
            check1 = list(InputCheck(temp1, 'قیمت یک واحد قطعه '))
            check2 = list(InputCheck(temp2, "تعداد واحد مصرفي "))
            # check3=InputCheck(temp3,"نام کالای مصرفی ")
            if check1[0] == 0:
                self.flag7 = 1
                messagebox.showinfo("Message", check1[1])
                break
            else:
                self.unit_price.append(check1[1])
            if check2[0] == 0:
                self.flag7 = 1
                messagebox.showinfo("Message", check2[1])
                break
            else:
                self.consumption.append(check2[1])

        b1 = self.h1.get()
        b2 = self.h2.get()
        b3 = self.h3.get()
        check4 = list(InputCheck(b1, 'هزینه واحد زمان عملیات '))
        check5 = list(InputCheck(b2, 'زمان عملیات '))
        check6 = list(InputCheck(b3, 'درصد هزینه سربار '))
        

    def new_row(self):
        global i
        i += 1
        self.creat_e()

    def myfunction(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"), width=430, height=115)

    def destroy_all(self):
        global flag5
        flag5 = 1
        self.root.destroy()

    def creat_window1(self):
        self.wc_number += 1
        self.m1 = "اطلاعات ايستگاه شماره {}".format(self.wc_number)
        self.root = Tk()
        self.root.title("محاسبه قیمت نیمه ساخته")
        self.root.minsize(width=100, height=420)
        self.root.resizable(width=False, height=False)

        self.var1 = StringVar()
        self.var1.set(self.m1)

        self.myframe = Frame(self.root, relief=GROOVE, bd=1)
        self.myframe.place(x=1, y=290)
        # myframe2 = Frame(root, relief=GROOVE, bd=1)
        # myframe2.grid(row=0, column=0)
        self.canvas = Canvas(self.myframe)
        self.frame = Frame(self.canvas)
        self.myscrollbar = Scrollbar(self.myframe, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.myscrollbar.set)

        self.myscrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", expand=True, fill="x")
        self.canvas.create_window((0, 0), window=self.frame, anchor='nw')
        self.frame.bind("<Configure>", self.myfunction)

        Label(self.root, text=self.var1.get(), font="bnazanin 14 bold").place(x=200, y=15)
        Label(self.root, text=' ', font="bnazanin 14 bold", pady=15).grid(row=0, column=0, padx=80)

        Label(self.root, text="قیمت یک واحد", font="bnazanin 9 bold").grid(row=4, column=0)
        Label(self.root, text="تعداد واحد مصرفي", font="bnazanin 9 bold", justify=CENTER).grid(row=4, column=1,
                                                                                               sticky=W,
                                                                                               padx=(6, 0))
        Label(self.root, text="نام کالای مصرفی", font="bnazanin 9 bold").grid(row=4, column=2)
        Label(self.root, text="زمان انجام عملیات \nدر ایستگاه", font="bnazanin 9 bold").grid(row=1, column=3)
        Label(self.root, text="هزینه واحد زمان عملیات", font="bnazanin 9 bold", pady=8).grid(row=1, column=1, pady=15)
        Label(self.root, text="درصد هزینه سربار\n واحد زمان", font="bnazanin 9 bold").grid(row=2, column=1)
        Label(self.root, text="________________________________________________", font="bnazanin 9").place(x=30, y=184)

        self.h1 = Entry(self.root, width=10)
        # h1.insert(10, 1)
        self.h1.grid(row=1, column=0)
        self.h2 = Entry(self.root, width=10)
        # h2.insert(10, 1)
        self.h2.grid(row=1, column=2, padx=25)
        self.h3 = Entry(self.root, width=10)
        self.h3.insert(10, 18)
        self.h3.grid(row=2, column=0)

        self.creat_e()

        pad_x2 = 27
        Button(self.root, text="ایجاد ردیف جدید", command=self.new_row, font="bnazanin 10 bold",
               background='#eaeaea').grid(row=2,
                                          column=3,
                                          padx=pad_x2,
                                          pady=10)
        Button(self.root, text="ثبت و ادامه", command=self.save_fields, font="bnazanin 10 bold",
               background='#eaeaea').grid(row=3,
                                          column=3,
                                          padx=pad_x2,
                                          pady=10)
        Button(self.root, text="خروج", command=self.destroy_all, font="bnazanin 10 bold", background='#eaeaea').grid(
            row=4,
            column=3,
            padx=pad_x2,
            pady=10)

        SetScreen(self.root, 7)
        self.root.mainloop()


class Result:

    def __init__(self):
        self.master = None
        self.main_with = None
        self.f1 = None
        self.f2 = None
        self.f3 = None
        self.k1 = None
        self.creat_window()

    def pl0(self):
        global wc_price
        labels = []
        for r in range(0, len(wc_price)):
            b = r + 1
            # y='cost of the \nwork center {}'.format(str(b))
            y = 'ﻩﺎﮕﺘﺴﯾﺍ ﻪﻨﯾﺰﻫ \n ﻩﺭﺎﻤﺷ : {}'.format(str(b))
            labels.append(y)
        # title ="Final price pie chart \n \n Work Center number separation"
        title = 'ﯽﯾﺎﻬﻧ ﺖﻤﯿﻗ ﺭﺍﺩﻮﻤﻧ \n ﻩﺎﮕﺘﺴﯾﺍ ﻩﺭﺎﻤﺷ ﻚﯿﻜﻔﺗ ﺎﺑ'
        sizes = wc_price
        Plot(sizes, labels, title, flag=0)

    def pl1(self):
        global total_wc_workpiece_price, total_wc_time_price, total_wc_overhead_price
        sizes = [total_wc_workpiece_price, total_wc_time_price, total_wc_overhead_price]
        #labels = ['Total cost of used \npiece of all WC', 'Total cost of work \ntime of all WC',
        #          'Total overhead \ncost of all WC']
      
    def pl2(self):
        global all_work_center
        labels = []
        sizes = []
        for r in range(0, len(all_work_center)):
            sizes.append(all_work_center[r][8])
            number = r + 1
            # y='cost of the \nwork center {}'.format(str(b))
            y = 'ﺭﺎﺑﺮﺳ ﻪﻨﯾﺰﻫ \n ﻩﺭﺎﻤﺷ ﻩﺎﮕﺘﺴﯾﺍ : {}'.format(str(number))
            labels.append(y)
        # title ="Final price pie chart \n \n Work Center number separation"
        title = 'ﺭﺎﺑﺮﺳ ﻪﻨﯾﺰﻫ ﻪﺴﯾﺎﻘﻣ ﺭﺍﺩﻮﻤﻧ \n ﻒﻠﺘﺨﻣ ىﺎﻫ ﻩﺎﮕﺘﺴﯾﺍ'
        Plot(sizes, labels, title)

    def pl3(self):
        global all_work_center
        u = self.k1.get()
        check = list(InputCheck(u, 'شماره ایستگاه '))
        if check[0] == 0:
            messagebox.showinfo("Message", check[1])
        else:
            if check[1] == round(check[1]):
                if check[1] <= len(all_work_center):
                    u = int(check[1])
                    sizes = all_work_center[u - 1][6:9]
                    # labels = ['Total cost of \nparts', 'Cost of operation \ntime', 'Overhead cost']
                    labels = ['ﻪﻨﯾﺰﻫ ﻉﻮﻤﺠﻣ \n ﺕﺎﻌﻄﻗ', 'ﻥﺎﻣﺯ ﻪﻨﯾﺰﻫ \n ﺭﺎﻛ ﻡﺎﺠﻧﺍ', 'ﺭﺎﺑﺮﺳ ﻪﻨﯾﺰﻫ']
                    # title = "Pie chart for \n Work center {}".format(u)
                    title = '{} ﻩﺭﺎﻤﺷ ﻩﺎﮕﺘﺴﯾﺍ ﺭﺍﺩﻮﻤﻧ '.format(u)
                    Plot(sizes, labels, title)
                else:
                    messagebox.showinfo("Message", "ایستگاهی با این شماره تعریف نشده است")
            else:
                messagebox.showinfo("Message", "شماره ايستگاه باید عدد صحیح باشد")

    def creat_window(self):
        global final_price, total_wc_workpiece_price, total_wc_time_price, total_wc_overhead_price, workpiece_name
        self.master = Tk()
        self.master.title("محاسبه قیمت نیمه ساخته")
        self.master.minsize(100, 150)
        self.main_with = 450
        self.master.resizable(width=False, height=False)

        if workpiece_name == '':
            workpiece_name = 'نیمه ساخته'
        self.f1 = Frame(self.master)
        self.f1.grid(row=0, column=0, sticky="nsew", pady=22)
        Label(self.f1, text=' ', font="bnazanin 14 bold").grid(row=0, column=0, padx=42)
        Label(self.f1, text=workpiece_name, font="bnazanin 14 bold").grid(row=0, column=1)
        Label(self.f1, text="گزارش هزینه های تولید ", font="bnazanin 14 bold").grid(row=0, column=2)
        self.master.grid_columnconfigure(0, minsize=self.main_with)

        Label(self.master, text='قیمت تمام شده نیمه ساخته : {} تومان'.format(final_price), font="bnazanin 9 bold",
              justify=RIGHT, padx=20).grid(row=1, column=0, pady=15, sticky=E)
              font="bnazanin 9 bold", justify=RIGHT, padx=20).grid(row=3, column=0, pady=15, sticky=E)
        Label(self.master, text="مجموع هزینه سربار تمام ایستگاه ها :  {} تومان".format(total_wc_overhead_price),
              font="bnazanin 9 bold", justify=RIGHT, padx=20).grid(row=4, column=0, pady=15, sticky=E)
        # Label (master, text="نام کالای مصرفی").place(x=300,y=50)

        self.f2 = Frame(self.master)
        self.f2.grid(row=5, column=0, sticky="nsew", pady=22)
        Label(self.f2, text="نمودار قیمت نهایی\n (با تفکیک شماره ایستگاه)").grid(row=0, column=2, padx=(20, 15))
        Label(self.f2, text="نمودار قیمت نهایی\n (با تفکیک نوع هزینه)").grid(row=0, column=1, padx=(20, 0))
        Label(self.f2, text="نمودار مقایسه هزینه سربار\nایستگاه های مختلف").grid(row=0, column=0, padx=(20, 0))

        Button(self.f2, text="رسم نمودار", command=self.pl0, padx=4, pady=2, font="bnazanin 10 bold",
               background='#eaeaea').grid(
            row=1, column=2, pady=(20, 5), padx=(5, 0))
        Button(self.f2, text="رسم نمودار", command=self.pl1, padx=4, pady=2, font="bnazanin 10 bold",
               background='#eaeaea').grid(
            row=1, column=1, pady=(20, 5), padx=(22, 0))
        Button(self.f2, text="رسم نمودار", command=self.pl2, padx=4, pady=2, font="bnazanin 10 bold",
               background='#eaeaea').grid(
            row=1, column=0, pady=(20, 5), padx=(20, 0))

        Label(self.master, text="________________________________________________________").grid(row=6, column=0)
        Label(self.master, text=" ترسیم نمودار یک ایستگاه خاص", font="bnazanin 14 bold").grid(row=7, column=0,
                                                                                              pady=(20, 5))

        self.f3 = Frame(self.master)
        self.f3.grid(row=8, column=0, sticky="nsew")
        Label(self.f3, text=" : شماره ایستگاه مورد نظر").grid(row=0, column=2, padx=(30, 0), pady=(20, 5))
        self.k1 = Entry(self.f3, width=7)
        self.k1.insert(10, 1)
        self.k1.grid(row=0, column=1, padx=(50, 0), pady=(20, 5))

        Button(self.f3, text="رسم نمودار", command=self.pl3, padx=6, pady=2, font="bnazanin 10 bold",
               background='#eaeaea').grid(
            row=0, column=0, padx=(40, 0), pady=(20, 5))

        Button(self.master, text="خروج", command=self.master.destroy, padx=6, pady=2, font="bnazanin 10 bold",
               background='#eaeaea').grid(row=9, column=0, pady=(20, 30))
        # Button(master, text="ثبت", command='').place(x=50, y=600)

        SetScreen(self.master, 7, flag_n=1)
        mainloop()


class Plot:

    def __init__(self, sizes, labels, title, flag =1):
        self.sizes = sizes
        self.labels = labels
        self.title = title
        self.flag = flag
        self.creat_plot()

    def creat_plot(self):
        persian_labels = self.labels
        persian_title = self.title
        explode = []
        explode = [0] * len(self.sizes)
        # if flag == 1:
        #    explode = [0,0,0.4]

        plt.close()
        pie = plt.pie(self.sizes, autopct='%1.1f%%', explode=explode)
        plt.title(persian_title)
        plt.axis('equal')

        labels2 = []
        for i in range(len(self.sizes)):
            temp = '{} = {}'.format(persian_labels[i], self.sizes[i])
            labels2.append(temp)
        plt.legend(bbox_to_anchor=(0.60, 0.5, 0.52, 0.65), labels=labels2, prop=dict(size=8))

        plt.show()


SelectInputMethod()

if select_method == 1:
    WcNumber()
    all_work_center = [0] * n
    for g in range(0, n):
        if flag1 == 1 and flag5 == 0:
            all_work_center[g] = list(WorkCenter(g))
            i = 0
    if flag1 == 1 and flag5 == 0:
        flag2 = 1

if select_method == 2:
    ExcelDataEntry()

if flag2 != 0:
    wc_price = []
    for g in range(0, len(all_work_center)):
        wc_price_temp = 0
        unit_workpiece_price = all_work_center[g][0]
        amount_workpiece = all_work_center[g][1]
        unit_time_price = round(float(all_work_center[g][3]), 2)
        wc_time = round(float(all_work_center[g][4]), 2)
        percentage_overhead = round(float(all_work_center[g][5]), 2)
        all_workpiece_price = 0
        for y in range(len(unit_workpiece_price)):
            unit_workpiece_price[y] = round(float(unit_workpiece_price[y]), 2)
            amount_workpiece[y] = round(float(amount_workpiece[y]), 2)
            all_workpiece_price = all_workpiece_price + (unit_workpiece_price[y] * amount_workpiece[y])

        all_work_center[g].append(round(float(all_workpiece_price), 2))
        wc_total_time_price = wc_time * unit_time_price
        all_work_center[g].append(round(float(wc_total_time_price), 2))
        overhead_price = (percentage_overhead * wc_total_time_price) / 100
        all_work_center[g].append(round(float(overhead_price), 2))
    for k in range(0, len(wc_price)):
        final_price += wc_price[k]
    final_price = round(float(final_price), 2)
    total_wc_workpiece_price = 0
    total_wc_time_price = 0
    total_wc_overhead_price = 0
    for k in range(0, len(all_work_center)):
        total_wc_workpiece_price += all_work_center[k][6]
        total_wc_time_price += all_work_center[k][7]
        total_wc_overhead_price += all_work_center[k][8]
    total_wc_workpiece_price = round(float(total_wc_workpiece_price), 2)
    total_wc_time_price = round(float(total_wc_time_price), 2)
    total_wc_overhead_price = round(float(total_wc_overhead_price), 2)
    Result()
