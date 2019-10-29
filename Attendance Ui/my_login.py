"""
This software is developed by Soumil Shah under ECT labs

soumil shah

Bachelor in Electronic Engineering
Master in Electrical Engineering
Master in Computer Engineering

"""


try:

    import pyttsx3
    import tkinter as tk
    from tkinter import messagebox
    from UI_Read_RFID import *
    from ui_add_db import *
    import datetime
    import time
    from xlwt import Workbook
    from xlwt import easyxf
    from twilio.rest import Client
    from tkinter import PhotoImage
    from mygoogle import *
    from namegoogle import  *
    import threading
except:
    print('library not found ')


running = True  # Global flag
idx = 0  # loop index
Freq = 2500
Dur = 150
LARGE_FONT= ("Verdana", 12)

global l
l=[]

global flag
flag="Status"


class Attendance(tk.Tk):

    def __init__(self, *args, **kwargs):

        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)

        container.pack(side="top", fill="both", expand = True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, PageOne, PageTwo,PageThree,SmsPage,Cloud):

            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)

    def show_frame(self, cont):

        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)

        # row 0--------------------------------------

        self.logo = tk.PhotoImage(file='ub.gif')
        w1=tk.Label(self,image=self.logo)
        w1.grid(row=0,column=0, sticky="NSEW",padx=10,pady=10)

        l_title=tk.Label(self, text="IoT Based Attendance System",
                         font="Arial,12")
        l_title.grid(row=0,column=1,columnspan=2, sticky="NSEW",padx=10,pady=10)

        clock = tk.Label(self, font=('times', 18, 'bold'),fg="RED")
        clock.grid(row=0,column=3, sticky="NSNESWSE",padx=8)

        def tick():
            time2=time.strftime('%H:%M:%S')
            clock.config(text=time2)
            clock.after(200,tick)
        tick()

        #----------------------------------------------------
        # row 1

        label_username = tk.Label(self, text="Username")
        label_password = tk.Label(self, text="Password")

        entry_username = tk.Entry(self)
        entry_password = tk.Entry(self, show="*")

        l_1=tk.Label(self)
        l_1.grid(row=1, column=0, sticky='NSEW',padx=10,pady=10)

        label_username.grid(row=3, column=1, sticky='NSEW',padx=10,pady=10)
        entry_username.grid(row=3, column=2,sticky='NSEW',padx=10,pady=10)
        label_password.grid(row=4, column=1, sticky='NSEW',padx=10,pady=10)
        entry_password.grid(row=4, column=2,sticky='NSEW',padx=10,pady=10)

        checkbox = tk.Checkbutton(self, text="Keep me logged in")
        checkbox.grid(row=5, column=2,sticky='NSEW',padx=10,pady=10)

        logbtn = tk.Button(self, text="Login", bg="BlACK",fg="White",
                           command=lambda: login_btn_clicked(),height=2, width=8)
        logbtn.grid(row=6, column=2,sticky='NSEW', padx=10, pady=10)


        def login_btn_clicked():
            try:

                # print("Clicked")
                username = entry_username.get()
                password = entry_password.get()

            
                if len(username) and len(password) > 2:
                    # print(username, password)

                    if username == "admin" and password == "admin":
                        mm_message=""" Welcome Admin
                            You can Add Student or Delete Student or View Existing Student 
                            Enter Com port and Baud Rate and press scan RFID  Button to assign a Tag
                                         """

                        t=threading.Thread(target=my_message,args=(mm_message,))
                        t.start()

                        controller.show_frame(PageOne)

                    else:
                        t=threading.Thread(target=my_message,args=('Invalid Credentials please try again ',))
                        t.start()
                        messagebox.showinfo(self,"Invalid Credentials please try again ")
            except:
                print('Cannot Execute Login function')
                messagebox.showinfo(self,"Cannot Execute Login function")


        def my_message(my_message):
            try:

                engine = pyttsx3.init()
                rate = engine.getProperty('rate')
                engine.setProperty('rate', rate-50)
                engine.say('{}'.format(my_message))
                engine.runAndWait()
                #rate = engine.getProperty('rate')
            except:
                print('Faield to execute my_message function ! ')

        t=threading.Thread(target=my_message,args=('I am Smart Bot i will assist you regarding student attendance please enter your credentials',))
        t.start()


class PageOne(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)



    # ---------------------------------------------------------------------

        clock = tk.Label(self, font=('times', 18, 'bold'), bg='green',fg="white")
        clock.grid(row=0,column=2, sticky="NSNESWSE",padx=8,pady=8)

        def tick():
            time2=time.strftime('%H:%M:%S')
            clock.config(text=time2)
            clock.after(200,tick)
        tick()


        # -----------------------------------------------------

        label = tk.Label(self, text="RFID Attendance System!", font="Arial,16",bg="black",fg="White")
        label.grid(row=0, column=0,columnspan=2,padx=8,pady=8, sticky="NSNESWSE")

        # --------------------------------------------------
        """
        Only Define First Name, middle Name, Last Name
        """

        l_name=tk.Label(self, text=" First Name ")
        l_name.grid(row=1, column=0, sticky="NSNESWSE",padx=10,pady=10)

        e_name=tk.Entry(self)
        e_name.grid(row=2, column=0 ,sticky="NSNESWSE", padx=10, pady=10)

        m_name=tk.Label(self, text=" Middle Name ")
        m_name.grid(row=1, column=1, sticky="NSNESWSE",padx=10,pady=10)

        e_name_m=tk.Entry(self)
        e_name_m.grid(row=2, column=1 ,sticky="NSNESWSE", padx=10, pady=10)

        m_name=tk.Label(self, text=" Last Name ")
        m_name.grid(row=1, column=2, sticky="NSNESWSE",padx=10,pady=10)

        e_name_l=tk.Entry(self)
        e_name_l.grid(row=2, column=2, sticky="NSNESWSE", padx=10, pady=10)

        # ------------------------------------------------------
        """
        we are working with row 3 is label for email, phone , id
        row 4 is for entry widget 
        
        """

        l_email = tk.Label(self,text="Enter Email")
        l_email.grid(row=3, column=0, sticky="NSNESWSE", padx=10, pady=10)

        e_email = tk.Entry(self)
        e_email.grid(row=4, column=0, sticky="NSNESWSE", padx=10, pady=10)

        l_phone = tk.Label(self,text="Enter Phone Number ")
        l_phone.grid(row=3, column=1, sticky="NSNESWSE", padx=10, pady=10)

        e_phone = tk.Entry(self)
        e_phone.grid(row=4, column=1, sticky="NSNESWSE", padx=10, pady=10)

        l_id = tk.Label(self, text="Enter Id")
        l_id.grid(row=3, column=2,sticky="NSNESWSE")

        e_id = tk.Entry(self)
        e_id.grid(row=4, column=2, sticky="NSNESWSE", padx=10, pady=10)

        # --------------------------------------------------------------------------

        """
        This is row 5 is for label of gender, scan rfid,press button
        row 6 is for all entry widget 
        """

        l_gender=tk.Label(self,text="Gender")
        l_gender.grid(row=5, column=0, sticky="NSNESWSE", padx=10, pady=10)

        e_gender = tk.Entry(self,)
        e_gender.grid(row=6, column=0, sticky="NSNESWSE", padx=10, pady=10)

        l_rfid = tk.Label(self, text="Enter/Scan RFID ")
        l_rfid.grid(row=5, column=1, sticky="NSNESWSE")

        e_rfid = tk.Entry(self,)
        e_rfid.grid(row=6, column=1, sticky="NSNESWSE", padx=10, pady=10)

        l_scan_m = tk.Label(self, text="Click Button to Scan Card")
        l_scan_m.grid(row=5, column=2, sticky="NSNESWSE")

        b_scan=tk.Button(self,text="Scan RFID", background="Blue", fg="White",
                         command=lambda: scan())
        b_scan.grid(row=6,column=2, sticky="NSNESWSE", padx=10, pady=10)

        # --------------------------------------------------------------------------
        """
        we are working with row 7 which has 
        add user, view, delete user 
        """

        b_add_user=tk.Button(self, text="Add User",command=lambda:add_database(),
                             bg="Black",fg="White")
        b_add_user.grid(row=7, column=0, sticky="NSNESWSE", padx=5, pady=10)

        b_deleteuser=tk.Button(self,text="Delete User", command=lambda: delete_database(),
                               bg="Black",fg="White")
        b_deleteuser.grid(row=7, column=1, sticky="NSNESWSE", padx=5, pady=10)

        b_show=tk.Button(self,text="Show User",command=lambda: show_user(),
                         bg="Black",fg="White")
        b_show.grid(row=7, column=2, sticky= "NSNESWSE", padx=5, pady=10)

        # -------------------------------------------------------------------
        """
        we are working with row number 8 for text display 
        """

        scrollbar_y = tk.Scrollbar(self)
        scrollbar_y.grid(row=8, column=3)

        show_1=tk.Text(self,height=8, width=35, yscrollcommand=scrollbar_y.set,
                       bg="Grey",fg="White")
        show_1.grid(row=8, column=0,rowspan=3,columnspan=3,sticky="NSEW")

        #-------------------------------------------------------------------

        serial_label=tk.Label(self,text="Enter Correct COM Port and Baud Rate")
        serial_label.grid(row=17, column=0, padx=10, pady=10,sticky="NSNESWSE")

        serial_e=tk.Entry(self)
        serial_e.grid(row=17, column=1, padx=10, pady=10,sticky="NSNESWSE")

        e_baud=tk.Entry(self)
        e_baud.grid(row=17, column=2, padx=10, pady=10,sticky="NSNESWSE")

        status=tk.Label(self,text="Status",bg="grey",fg="White")
        status.grid(row=18, column=1,columnspan=2, padx=10, pady=10,sticky="NSNESWSE")



        # -------------------------------------------------------------------

        button1 = tk.Button(self, text="Back to Home",
                            command=lambda: controller.show_frame(StartPage),
                            bg="Black",fg="White")
        button1.grid(row=19, column=0, padx=5, pady=5,sticky="NSNESWSE")

        button2 = tk.Button(self, text=" Attendance ",
                    command=lambda: controller.show_frame(PageTwo),
                            bg="Black",fg="White")
        button2.grid(row=19,column=1,padx=5, pady=5,sticky="NSNESWSE")

        clear_text = tk.Button(self, text=" Clear Text ",
                            command=lambda:my_text(),
                               bg="Black",fg="White")
        clear_text.grid(row=19,column=2,padx=5, pady=5,sticky="NSNESWSE")

        # -----------------------------------------------------------------------------
        def add_database():
            try:
                # -----------------------------
                # get name
                e_name_v= e_name.get()
                e_name_m_v= e_name_m.get()
                e_name_l_v= e_name_l.get()

                e_email_v= e_email.get()
                e_id_v= e_id.get()
                e_phone_v= e_phone.get()

                e_rfid_v= e_rfid.get()
                e_gender_v= e_gender.get()

                if(len(e_name_v)and len(e_name_m_v) and len(e_name_l_v) and len(e_email_v)
                 and len(e_id_v) and len(e_phone_v) and len(e_rfid_v) and
                len(e_gender_v)) >=1:
                    add_data_database(e_name_v, e_name_m_v,e_name_l_v,
                                      e_email_v, e_id_v, e_phone_v,
                                      e_rfid_v, e_gender_v)

                    messagebox.showinfo("User Added", " Name{} Rfid Tag {}"
                                        .format(e_name_v, e_rfid_v))
                    p=threading.Thread(target=my_speak,args=('{} was Added to Database'.format(e_name_v),))
                    p.start()
                else:
                    p=threading.Thread(target=my_speak,args=('Please Enter all feilds Correctly',))
                    p.start()

                    messagebox.showinfo(self,"Please Fill all data")


            except:
                print('CANNOT ADD  TO DATABASE')
                messagebox.showinfo(self,"CANNOT ADD  TO DATABASE")


        def scan():
            try:
                # check the status if connected if not message box pop not connected
                serial_e_v=serial_e.get()
                e_baud_v=e_baud.get()

                if len(serial_e_v) > 2 and len(e_baud_v) > 2:

                    p=threading.Thread(target=my_speak,args=('Place Your Card on Reader',))
                    p.start()


                    temp=read_card(serial_e_v,e_baud_v)

                    e_rfid.insert(0,temp)
                    status=tk.Label(self,text="Status",bg="Green",fg="White")
                    status.grid(row=18, column=1,columnspan=2, padx=10, pady=10,sticky="NSNESWSE")


                else:
                    p=threading.Thread(target=my_speak,args=('Please Enter Correct Com Port',))
                    p.start()

                    status=tk.Label(self,text="Status",bg="RED",fg="White")
                    status.grid(row=18, column=1,columnspan=2, padx=10, pady=10,sticky="NSNESWSE")

            except:
                print('cannot scan rfid card ')
                messagebox.showinfo(self,"cannot scan rfid card ")
    # --------------------------------------------------------------------------

        def show_user():
            try:

                p=threading.Thread(target=my_speak,args=('Please Wait ',))
                p.start()

                l=[]
                conn=sqlite3.connect('UI_User.db')
                c=conn.cursor()
                c.execute('SELECT * FROM my_student ')
                for x in c.fetchall():
                    # print(" ", x)
                    l.append(x)
                    show_1.insert(tk.END,x)
                    show_1.insert(tk.END,"\n")
            except:
                print('Could not Read from Database ')
                messagebox.showinfo(self,"Failed to execute show_user function ")
        # ---------------------------------------------------------------

        def delete_database():
            try:
                e_name_v = e_name.get()
                e_id_v= e_id.get()
                e_rfid_v = e_rfid.get()

                if len(e_name_v) >= 2:
                    conn=sqlite3.connect('UI_User.db')
                    c=conn.cursor()
                    c.execute("DELETE FROM  my_student WHERE fname=?    ", (e_name_v,))
                    conn.commit()
                    messagebox.showinfo("User Deleted Name{}" .format(e_name_v))

                    p=threading.Thread(target=my_speak,args=('{} was Deleted from DataBase'.format(e_name_v),))
                    p.start()


                else:
                    p=threading.Thread(target=my_speak,args=('Please Enter First Name ',))
                    p.start()
                    messagebox.showinfo(self,"Please Enter First Name")

            except:
                print ('cannot database')
                messagebox.showinfo(self,"Cannot Delete ")

            # --------------------------------------------------------------------
        def my_text():
            try:
                show_1.delete('1.0',tk.END)
            except:
                print("Faailed to exectute my_text ")

        # ---------------------------------------------

        def my_speak(my_message):
            try :
                engine = pyttsx3.init()
                rate = engine.getProperty('rate')
                engine.setProperty('rate', rate-20)
                engine.say('{}'.format(my_message))
                engine.runAndWait()
                #rate = engine.getProperty('rate')
            except:
                print('cannot execute my_speak Function')



class PageTwo(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        # ---------------------------------------------------------------------

        clock = tk.Label(self, font=('times', 18, 'bold'), bg='green',fg="white")
        clock.grid(row=0,column=2, sticky="NSEW",padx=10,pady=10)

        def tick():
            time2=time.strftime('%H:%M:%S')
            clock.config(text=time2)
            clock.after(200,tick)
        tick()
        # ------------------------------------------------------------------------------------------

        label = tk.Label(self, text="RFID Attendance System!", font="Arial,16",bg="black",fg="White")
        label.grid(row=0, column=0,padx=10,pady=10, sticky="NSEW")

        # --------------------------------------------------------------------------------------------

        """
        Only Define First Name, middle Name, Last Name
        """

        l_name=tk.Label(self, text=" First Name ")
        l_name.grid(row=1, column=0, sticky="NSNESWSE",padx=10,pady=10)

        e_name=tk.Entry(self)
        e_name.grid(row=2, column=0 ,sticky="NSNESWSE", padx=10, pady=10)

        m_name=tk.Label(self, text=" Middle Name ")
        m_name.grid(row=1, column=1, sticky="NSNESWSE",padx=10,pady=10)

        e_name_m=tk.Entry(self)
        e_name_m.grid(row=2, column=1 ,sticky="NSNESWSE", padx=10, pady=10)

        m_name=tk.Label(self, text=" Last Name ")
        m_name.grid(row=1, column=2, sticky="NSNESWSE",padx=10,pady=10)

        e_name_l=tk.Entry(self)
        e_name_l.grid(row=2, column=2, sticky="NSNESWSE", padx=10, pady=10)

        # ------------------------------------------------------
        """
        we are working with row 3 is label for email, phone , id
        row 4 is for entry widget 
        
        """

        l_email = tk.Label(self,text="Enter Email")
        l_email.grid(row=3, column=0, sticky="NSNESWSE", padx=10, pady=10)

        e_email = tk.Entry(self)
        e_email.grid(row=4, column=0, sticky="NSNESWSE", padx=10, pady=10)

        l_phone = tk.Label(self,text="Enter Phone Number ")
        l_phone.grid(row=3, column=1, sticky="NSNESWSE", padx=10, pady=10)

        e_phone = tk.Entry(self)
        e_phone.grid(row=4, column=1, sticky="NSNESWSE", padx=10, pady=10)

        l_id = tk.Label(self, text="Enter Id")
        l_id.grid(row=3, column=2,sticky="NSNESWSE")

        e_id = tk.Entry(self)
        e_id.grid(row=4, column=2, sticky="NSNESWSE", padx=10, pady=10)

        # --------------------------------------------------------------------------

        """
        This is row 5 is for label of gender, scan rfid,press button
        row 6 is for all entry widget 
        """

        l_gender=tk.Label(self,text="Gender")
        l_gender.grid(row=5, column=0, sticky="NSNESWSE", padx=10, pady=10)

        e_gender = tk.Entry(self,)
        e_gender.grid(row=6, column=0, sticky="NSNESWSE", padx=10, pady=10)

        l_rfid = tk.Label(self, text="Enter/Scan RFID ")
        l_rfid.grid(row=5, column=1, sticky="NSNESWSE")

        e_rfid = tk.Entry(self,)
        e_rfid.grid(row=6, column=1, sticky="NSNESWSE", padx=10, pady=10)

        l_scan_m = tk.Label(self, text="Time")
        l_scan_m.grid(row=5, column=2, sticky="NSNESWSE")

        e_time=tk.Entry(self,text="Scan RFID",)
        e_time.grid(row=6,column=2, sticky="NSNESWSE", padx=10, pady=10)

        # --------------------------------------------------------------------------
        """
        we are working with row number 8 for text display 
        """

        m_label=tk.Label(self,text="Swipe Your card for Attendance",bg="Black",fg="WHITE",
                         font=("Arial",12))
        m_label.grid(row=7,column=0,columnspan=3,sticky="NSNESWSE",padx=10,pady=10)

        scrollbar_y = tk.Scrollbar(self)
        scrollbar_y.grid(row=8, column=3,sticky="NSNESWSE")

        show_1=tk.Text(self,height=5, width=28, yscrollcommand=scrollbar_y.set,
                       bg="Grey",fg="White")
        show_1.grid(row=8, column=0,rowspan=3,columnspan=3,sticky="NSNESWSE")


        # ---------------------------------------------------------------------------

        serial_label=tk.Label(self,text="COMPort/BaudRate")
        serial_label.grid(row=16, column=0, padx=8, pady=8,sticky="NSNESWSE")

        serial_e=tk.Entry(self)
        serial_e.grid(row=16, column=1, padx=8, pady=8,sticky="NSNESWSE")

        e_baud=tk.Entry(self)
        e_baud.grid(row=16, column=2, padx=8, pady=8,sticky="NSNESWSE")

        b_connect=tk.Button(self,text="Connect",bg="Green",fg="White",command=lambda:connect() )
        b_connect.grid(row=17, column=0, padx=8, pady=8,sticky="NSNESWSE")

        b_disconnect=tk.Button(self,text="Disconnect",bg="Red",fg="White",command=lambda:disconnect())
        b_disconnect.grid(row=17, column=1, padx=8, pady=8,sticky="NSNESWSE")

        label_status=tk.Label(self,text="Status",bg="Grey",fg="White")
        label_status.grid(row=17, column=2, padx=8, pady=8,sticky="NSNESWSE")

        b_back = tk.Button(self, text="Back",
                            command=lambda: controller.show_frame(PageOne),bg="Black",fg="White")
        b_back.grid(row=19,column=1, padx=8, pady=8,sticky="NSNESWSE")



        b_third=tk.Button(self,text="View Record",bg="Black",fg="White",
                          command=lambda: controller.show_frame(PageThree))
        b_third.grid(row=19,column=2, padx=8, pady=8,sticky="NSNESWSE")


        def runner():
            try:


                label_status=tk.Label(self,text="Status",bg="GREEN",fg="White")
                label_status.grid(row=17, column=2, padx=8, pady=8,sticky="NSNESWSE")

                #RFID CODE
                global after_id
                global secs
                secs += 1

                if secs % 2 == 0:  # every other second

                    print("Reading Card")
                    serial_e_v=serial_e.get()
                    e_baud_v=e_baud.get()
                    temp=read_card(serial_e_v,e_baud_v)

                    l=[]

                    conn=sqlite3.connect('UI_User.db')
                    c=conn.cursor()
                    c.execute('SELECT * FROM my_student')

                    for x in c.fetchall():
                        l.append(x[6]) # Tag
                    # print("L",l)

                    if temp in l:
                        c.execute('SELECT * FROM my_student WHERE rfid=? ', (temp,))
                        data=data=c.fetchall()
                        print(data)

                        f_name = data[0][0]
                        m_name= data [0][1]
                        l_name= data [0][2]
                        e_email=data [0][3]
                        i_d    =data [0][4]
                        phone=  data [0][5]
                        g_gender=data[0][6]
                        my_time=time.strftime('%H:%M:%S')
                        my_date=now=datetime.datetime.today().strftime('%Y-%m-%d')
                        my_tag=temp


                        attendance_database(f_name, m_name, l_name, e_email, i_d, phone, g_gender, my_time, my_date,my_tag)

                        p11=threading.Thread(target=my_speak_pg2,args=('{} your attendance is recorded'.format(f_name),))
                        p11.start()
                after_id = self.after(1000, runner)  # check again in 1 second
            except:
                print('Could not execute Function runner ')
        def connect():
            # CONNECT COM PORT
            global secs
            secs = 0
            runner()  # start repeated checking


        def disconnect():
            # dISCONNECT cOM
            global after_id
            if after_id:
                self.after_cancel(after_id)
                after_id = None
                label_status=tk.Label(self,text="Status",bg="Red",fg="White")
                label_status.grid(row=17, column=2, padx=8, pady=8,sticky="NSNESWSE")
                p11=threading.Thread(target=my_speak_pg2,args=(' System Disconnected ',))
                p11.start()

        def attendance_database(f_name, m_name,
                                l_name, my_e_email,
                                i_d, phone,
                                g_gender, my_time,
                                my_date,my_tag):

            e_name.delete(0,tk.END)
            e_name_m.delete(0,tk.END)
            e_name_l.delete(0,tk.END)
            e_email.delete(0,tk.END)
            e_phone.delete(0,tk.END)
            e_id.delete(0,tk.END)
            e_gender.delete(0,tk.END)
            e_time.delete(0,tk.END)
            e_rfid.delete(0,tk.END)

            conn_my=sqlite3.connect('attendance.db')
            c_my = conn_my.cursor()

            c_my.execute(""" 
    CREATE TABLE IF NOT EXISTS my_student 
    (fname TEXT,mname TEXT,lname TEXT,email TEXT,id TEXT,phone TEXT,gender TEXT,time TEXT,date TEXT,rfid TEXT)""")


            c_my.execute("""INSERT INTO my_student 
    (fname, mname, lname, email, id, phone, gender, time, date, rfid) 
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", (f_name, m_name,l_name,my_e_email, i_d, phone,g_gender,my_time,my_date,my_tag))

            conn_my.commit()
            c_my.close()
            conn_my.close()

            e_name.insert(0,f_name)
            e_name_m.insert(0,m_name)
            e_name_l.insert(0,l_name)
            e_email.insert(0,my_e_email)
            e_phone.insert(0,phone)
            e_id.insert(0,i_d)
            e_gender.insert(0,g_gender)
            e_time.insert(0,my_time)
            e_rfid.insert(0,my_tag)
            show_1.insert(tk.END,"Attendance Marked {}".format(f_name))
            show_1.insert(tk.END,"\n")


        def my_speak_pg2(my_message):
            try:
                engine = pyttsx3.init()
                rate = engine.getProperty('rate')
                engine.setProperty('rate', rate)
                engine.say('{}'.format(my_message))
                engine.runAndWait()
                #rate = engine.getProperty('rate')
            except:
                print('Cannot execute my_speak function ')


class PageThree(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        # -------------------------------------------------------------------
        label = tk.Label(self, text="RFID Attendance System!", font="Arial,16",bg="black",fg="White")
        label.grid(row=0, column=0,padx=10,pady=10, sticky="NSEW")

        clock = tk.Label(self, font=('times', 18, 'bold'), bg='green',fg="white")
        clock.grid(row=0,column=2, sticky="NSEW",padx=10,pady=10)

        def tick():
            time2=time.strftime('%H:%M:%S')
            clock.config(text=time2)
            clock.after(200,tick)
        tick()
        #---------------------------------------------------------------------

        """
        from to date program
        """

        l_from_d=tk.Label(self,text="From Date")
        l_from_d.grid(row=1,column=0,padx=8,pady=8,sticky="NSEW")

        l_from_d=tk.Label(self,text="To Date")
        l_from_d.grid(row=1,column=1,padx=8,pady=8,sticky="NSEW")

        e_from_d=tk.Entry(self)
        e_from_d.grid(row=2,column=0,padx=8,pady=8,sticky="NSEW")

        e_from_t=tk.Entry(self)
        e_from_t.grid(row=2,column=1,padx=8,pady=8,sticky="NSEW")

        b_go=tk.Button(self,text="Download File",bg='Black',fg="White",
                       command=lambda :get_from_to_report())
        b_go.grid(row=2,column=2,padx=8,pady=8,sticky="NSEW")

        # -------------------------------------------------------
        """
        Report by current date 
        """

        l_current=tk.Label(self,text="Date")
        l_current.grid(row=3,column=0,padx=8,pady=8,sticky="NSEW")

        e_today=tk.Entry(self)
        e_today.grid(row=3,column=1,padx=8,pady=8,sticky="NSEW")

        b_go_t=tk.Button(self,text="Download File",bg="Black",fg="White",
                         command=lambda:today_report())
        b_go_t.grid(row=3,column=2,padx=8,pady=8,sticky="NSEW")

        # --------------------------------------------------------------------
        """
        for id report 
        """
        l_id=tk.Label(self,text="Enter Student Id")
        l_id.grid(row=4,column=0,padx=8,pady=8,sticky="NSEW")

        e_id=tk.Entry(self)
        e_id.grid(row=4,column=1,padx=8,pady=8,sticky="NSEW")

        b_id_go= tk.Button(self,text="Download File",bg="Black",fg="White",
                          command=lambda : generate_report())

        b_id_go.grid(row=4,column=2,padx=8,pady=8,sticky="NSEW")
        # --------------------------------------------------------
        """
        Report by name of students
        """

        l_name=tk.Label(self,text="Enter Student Name")
        l_name.grid(row=5,column=0,padx=8,pady=8,sticky="NSEW")

        e_name=tk.Entry(self)
        e_name.grid(row=5,column=1,padx=8,pady=8,sticky="NSEW")

        b_name_go=tk.Button(self,text="Download File",bg="Black",fg="White",
                            command=lambda:name_report())
        b_name_go.grid(row=5,column=2,padx=8,pady=8,sticky="NSEW")

        # ----------------------------------------------------------------
        l_present_student=tk.Label(self,text="Present Students :",
                                   bg="Black",fg="White",
                                   font=("Arial",12))
        l_present_student.grid(row=6,column=0,columnspan=3,padx=8,pady=8,sticky="NSEW")

        # ------------------------------------------------------------------------
        scrollbar_y = tk.Scrollbar(self)
        scrollbar_y.grid(row=7, column=3,sticky="NSNESWSE")

        show_1=tk.Text(self,height=8, width=28, yscrollcommand=scrollbar_y.set,
                       bg="Grey",fg="White")
        show_1.grid(row=7, column=0,columnspan=3,sticky="NSNESWSE")

        b_clear=tk.Button(self,text="Clear Screen",bg="Black",fg="WHITE",
                          command=lambda: clear_screen())
        b_clear.grid(row=8, column=0,columnspan=3,sticky="NSNESWSE")

        # ------------------------------------------------------------------

        b_dow_report=tk.Button(self,text="SMS",bg="Black",fg="White",
                               command=lambda :controller.show_frame(SmsPage))
        b_dow_report.grid(row=16,column=0,padx=8,pady=8,sticky="NSEW")

        b_dow_report_absent=tk.Button(self,text="Upload To  Cloud",bg="Black",fg="White",
                                      command=lambda:controller.show_frame(Cloud))
        b_dow_report_absent.grid(row=16,column=1,padx=8,pady=8,sticky="NSEW")

        b_dow_report_present=tk.Button(self,text="Download Raw Data",bg="Black",fg="White",
                                       command=lambda : raw_data_download())
        b_dow_report_present.grid(row=16,column=2,padx=8,pady=8,sticky="NSEW")

        button2 = tk.Button(self, text="Back",
                            command=lambda: controller.show_frame(PageOne),
                            bg="Blue",fg="white")
        button2.grid(row=17,column=0,columnspan=3,padx=8,pady=8,sticky="NSEW")


        # -------------------------------------------------------------------------------------

        def my_speak_pg33(my_message):

            engine = pyttsx3.init()
            rate = engine.getProperty('rate')
            engine.setProperty('rate', rate-20)
            engine.say('{}'.format(my_message))
            engine.runAndWait()
            #rate = engine.getProperty('rate')



        def google_spreadsheet():
            my_date_v=e_today.get() # get todays Date
            # fetch data from the database and display that day !

            if len(my_date_v) >= 5:  # check for valid data
                my_google(my_date_v)
            else:
                messagebox.showinfo(self,"Enter Date!")
                p_333=threading.Thread(target=my_speak_pg33,args=('Please Enter Date',))
                p_333.start()


        def raw_data_download():
            p_333=threading.Thread(target=my_speak_pg33,args=('PDownloading Raw Data',))
            p_333.start()

            conn=sqlite3.connect('attendance.db')
            c=conn.cursor()
            c.execute(""" SELECT * from my_student """)
            today=c.fetchall()

            wb_11=Workbook()
            sheet1_today_date_1 = wb_11.add_sheet('sheet 1')
            style1= easyxf('pattern: pattern solid, fore_colour red;')
            wb_11.save('Raw_data.xls')

            for data in range(0,len(today)):
                sheet1_today_date_1.write(data,0,today[data][0])
                sheet1_today_date_1.write(data,1,today[data][1])
                sheet1_today_date_1.write(data,2,today[data][2])
                sheet1_today_date_1.write(data,3,today[data][3])
                sheet1_today_date_1.write(data,4,today[data][4])
                sheet1_today_date_1.write(data,5,today[data][5])
                sheet1_today_date_1.write(data,6,today[data][6])
                sheet1_today_date_1.write(data,7,today[data][7])
                sheet1_today_date_1.write(data,8,today[data][8])


                wb_11.save('Raw_data.xls')
            messagebox.showinfo(self,"Report Created !")


        def today_report():

            my_date=e_today.get() # get todays Date
            # fetch data from the database and display that day !

            if len(my_date) >= 5:  # check for valid data



                conn=sqlite3.connect('attendance.db')
                c=conn.cursor()
                c.execute(""" SELECT fname,time,date from my_student WHERE date =?""",(my_date,))
                today=c.fetchall()

                # check if null data is yes print no data found
                if len(today) == 0 :
                    messagebox.showinfo(self,"No Data Found !")
                else:
                    # Go with Normal flow
                    show_1.insert(tk.END,"  Name        Time        Date")
                    show_1.insert(tk.END,"\n")

                    wb=Workbook()
                    sheet1_today_date = wb.add_sheet('sheet 1')
                    style1= easyxf('pattern: pattern solid, fore_colour red;')
                    wb.save('{}.xls'.format(my_date))
                    counter=0

                    for data in range(0,len(today)):

                        show_1.insert(tk.END," {} {} {}".format(today[data][0],today[data][1],today[data][2]))
                        show_1.insert(tk.END,"\n")
                        sheet1_today_date.write(data,0,today[data][0])
                        sheet1_today_date.write(data,1,today[data][1])
                        sheet1_today_date.write(data,2,today[data][2])
                        wb.save('{}.xls'.format(my_date))

                        counter=counter+1

                        if counter == len(today)-1:
                            sheet1_today_date.write(len(today)+1,0,"Total Student Attended",style1)
                            sheet1_today_date.write(len(today)+1,1,counter,style1)

                            p_333=threading.Thread(target=my_speak_pg33,args=('Report for {} Total {} Student were Present for lecture'.format(my_date,counter),))
                            p_333.start()

                            counter=0
                            messagebox.showinfo(self,"Report Created !")
            else:
                messagebox.showinfo(self,"Enter Date!")



        # -------------------------------------------------------------------------------

        def clear_screen():
            '''
            p_333=threading.Thread(target=my_speak_pg33,args=('Screen Cleared',))
            p_333.start()
            '''
            show_1.delete('1.0',tk.END)

# --------------------------------------------------------------------------------------------
        """
        Report based on Names
        
        """

        def name_report():
            try:


                e_name_v=e_name.get() # get todays Date

                if  len(e_name_v) >= 3:
                    # fetch data from the database and display that day !
                    conn=sqlite3.connect('attendance.db')
                    c=conn.cursor()
                    c.execute(""" SELECT fname,time,date from my_student WHERE fname =?""",(e_name_v,))
                    today=c.fetchall()
                    if len(today) == 0:
                        messagebox.showinfo(self,"No Record Found ! ")

                    else:

                        show_1.insert(tk.END,"  Name        Time        Date")
                        show_1.insert(tk.END,"\n")

                        wb=Workbook()
                        sheet1_name = wb.add_sheet('sheet 1')
                        style1= easyxf('pattern: pattern solid, fore_colour red;')
                        wb.save('{}.xls'.format(e_name_v))
                        sheet1_name.col(0).width=7000
                        sheet1_name.col(1).width=7000
                        sheet1_name.col(2).width=7000
                        counter=0


                        for data in range(0,len(today)):
                            show_1.insert(tk.END," {} {} {}".format(today[data][0],today[data][1],today[data][2]))
                            show_1.insert(tk.END,"\n")

                            print(today[data][0],today[data][1],today[data][2])

                            sheet1_name.write(data,0,today[data][0])
                            sheet1_name.write(data,1,today[data][1])
                            sheet1_name.write(data,2,today[data][2])
                            wb.save('{}.xls'.format(e_name_v))

                            counter=counter+1

                            if counter == len(today)-1:
                                sheet1_name.write(len(today)+1,0,"Total lecture Attended",style1)
                                sheet1_name.write(len(today)+1,1,counter,style1)


                                p_333=threading.Thread(target=my_speak_pg33,args=('Report for {} was downloaded on Computer and he attended {} lecture'.format(e_name_v,counter),))
                                p_333.start()
                                counter=0
                                messagebox.showinfo(self,"Report Created !")
                else:
                    messagebox.showinfo(self,"Enter Name !")
            except:
                print('failed to execute function name report ')

        # ----------------------------------------------------------------------

        """
        Genrates report based on Id
        
        """

        def generate_report():
            try:
                e_id_v=e_id.get() # get todays Date
                # fetch data from the database and display that day !

                # check if the user has entered any data
                if len(e_id_v) >= 3:

                    conn=sqlite3.connect('attendance.db')
                    c=conn.cursor()
                    c.execute(""" SELECT fname,time,date from my_student WHERE id =?""",(e_id_v,))
                    today=c.fetchall()

                    # write a logic to see if null data is received

                    if len(today) == 0:
                        messagebox.showinfo(self,"No Record Found !")
                    else:
                        wb=Workbook()
                        sheet1_id =wb.add_sheet('Sheet 1')
                        style1=easyxf('pattern: pattern solid, fore_colour red;')
                        wb.save('{}.xls'.format(e_id_v))

                        show_1.insert(tk.END,"  Name        Time        Date")
                        show_1.insert(tk.END,"\n")

                        counter=0

                        for data in range(0,len(today)):
                            counter=counter+1
                            show_1.insert(tk.END," {} {} {}".format(today[data][0],today[data][1],today[data][2]))
                            show_1.insert(tk.END,"\n")

                            print(today[data][0],today[data][1],today[data][2])

                            sheet1_id.write(data,0,today[data][0])
                            sheet1_id.write(data,1,today[data][1])
                            sheet1_id.write(data,2,today[data][2])

                            sheet1_id.col(0).width=7000
                            sheet1_id.col(1).width=7000
                            sheet1_id.col(2).width=7000
                            wb.save('{}.xls'.format(e_id_v))

                        p_333=threading.Thread(target=my_speak_pg33,args=('Report for ID no {} downloaded on computer and total lecture attended {} '.format(e_id_v,counter),))
                        p_333.start()
                        counter=0
                        messagebox.showinfo(self,"Report Created !")
                else:
                    messagebox.showinfo(self,"Please Enter Id !")
            except:
                print('failed to execute generate_report function')

    # ---------------------------------------------------------------------------------

        """
        Generates report in excel file when button clicked
        fro to date
        
        """

        def get_from_to_report():

            try:
                e_from_d_v=e_from_d.get()
                e_from_t_v=e_from_t.get()

                if (len(e_from_d_v) >= 6) and (len(e_from_t_v) >= 6):
                    # start the code
                    wb=Workbook()
                    sheet1 =wb.add_sheet('Sheet 1')
                    style1=easyxf('pattern: pattern solid, fore_colour red;')
                    wb.save('{}-{}.xls'.format(e_from_d_v,e_from_t_v))

                    counter=0

                    # fetch data from the database and display that day !
                    conn=sqlite3.connect('attendance.db')
                    c=conn.cursor()
                    c.execute(""" SELECT fname,time,date from my_student WHERE date BETWEEN ?  AND ?  """,(e_from_d_v,e_from_t_v,))
                    today=c.fetchall()

                    # Check if null data is received
                    if len(today) == 0:
                        messagebox.showinfo(self,"No Date Found !")
                    else:
                        show_1.insert(tk.END,"  Name        Time        Date")
                        show_1.insert(tk.END,"\n")

                        for data in range(0,len(today)):

                            show_1.insert(tk.END,"{} {} {}".format(today[data][0],today[data][1],today[data][2]))
                            show_1.insert(tk.END,"\n")

                            print(today[data][0],today[data][1],today[data][2])
                            sheet1.write(data,0,today[data][0])
                            sheet1.write(data,1,today[data][1])
                            sheet1.write(data,2,today[data][2])

                            sheet1.col(0).width=7000
                            sheet1.col(1).width=7000
                            sheet1.col(2).width=7000

                            wb.save('{}-{}.xls'.format(e_from_d_v,e_from_t_v))
                            counter=counter+1
                            if counter == len(today)-1:
                                sheet1.write(len(today)+1,0,"Total Student",style1)
                                sheet1.write(len(today)+1,1,counter,style1)
                                wb.save('{}-{}.xls'.format(e_from_d_v,e_from_t_v))
                                p_333=threading.Thread(target=my_speak_pg33,args=('from {} to {} total Number of student is {}'.format(e_from_d_v,e_from_t_v,counter),))
                                p_333.start()

                                counter=0
                                messagebox.showinfo(self,"Report Created !")
                else:
                    messagebox.showinfo(self,"Please Enter Date !")
            except:
                print('failed to execute get_from_to_report function ')




class SmsPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        # row 0--------------------------------------

        l_title=tk.Label(self, text="SMS Features",
                         font="Arial,12")
        l_title.grid(row=0,column=0,columnspan=2, sticky="NSEW",padx=10,pady=10)

        clock = tk.Label(self, font=('times', 18, 'bold'),fg="RED")
        clock.grid(row=0,column=3, sticky="NSNESWSE",padx=8)

        def tick():
            time2=time.strftime('%H:%M:%S')
            clock.config(text=time2)
            clock.after(200,tick)
        tick()
        # -------------------------------------------------------------


        l_account_ssid=tk.Label(self,text='Account SSID')
        l_account_ssid.grid(row=1,column=0, sticky="NSNESWSE",padx=14,pady=14)

        e_account_ssid=tk.Entry(self,show="*")
        e_account_ssid.grid(row=1,column=1, sticky="NSNESWSE",padx=14,pady=14)
        e_account_ssid.insert(tk.END,'ACbc2d4195f742b5641b3e295a4ffbc59e')

        l_Auth=tk.Label(self,text='Auth Token')
        l_Auth.grid(row=1,column=2, sticky="NSNESWSE",padx=14,pady=14)

        e_auth=tk.Entry(self,show='*')
        e_auth.grid(row=1,column=3, sticky="NSNESWSE",padx=14,pady=14)
        e_auth.insert(tk.END,'1404a6b960f0cde87812a4afb3bfc830')

        l_twilio_no=tk.Label(self,text='Twilio Number')
        l_twilio_no.grid(row=2,column=0, sticky="NSNESWSE",padx=14,pady=14)

        e_twilio_no=tk.Entry(self)
        e_twilio_no.grid(row=2,column=1, sticky="NSNESWSE",padx=14,pady=14)
        e_twilio_no.insert(tk.END,'+14243637976')

        l_to_number=tk.Label(self,text='To ')
        l_to_number.grid(row=2,column=2, sticky="NSNESWSE",padx=14,pady=14)

        e_to_number=tk.Entry(self,)
        e_to_number.grid(row=2,column=3, sticky="NSNESWSE",padx=14,pady=14)
        e_to_number.insert(tk.END,'+16462045957')

        l_date=tk.Label(self,text='Date')
        l_date.grid(row=3,column=0, sticky="NSNESWSE",padx=14,pady=14)

        e_date=tk.Entry(self)
        e_date.grid(row=3,column=1, sticky="NSNESWSE",padx=14,pady=14)

        l_date_format=tk.Label(self,text='YYYY-MM-DD')
        l_date_format.grid(row=3,column=2, sticky="NSNESWSE",padx=14,pady=14)

        b_go_date=tk.Button(self,text='Go',bg="Black",fg="white",
                            command=lambda :my_sms_date())
        b_go_date.grid(row=3,column=3, sticky="NSNESWSE",padx=14,pady=14)


        l_name=tk.Label(self,text='Name')
        l_name.grid(row=4,column=0, sticky="NSNESWSE",padx=14,pady=14)

        e_name=tk.Entry(self)
        e_name.grid(row=4,column=1, sticky="NSNESWSE",padx=14,pady=14)

        b_go_name=tk.Button(self,text='Go',bg="Black",fg="white",
                            command=lambda : name_report_sms())
        b_go_name.grid(row=4,column=3, sticky="NSNESWSE",padx=14,pady=14)

        def my_speak_sms(my_message):
            try:
                engine = pyttsx3.init()
                rate = engine.getProperty('rate')
                engine.setProperty('rate', rate)
                engine.say('{}'.format(my_message))
                engine.runAndWait()
                #rate = engine.getProperty('rate')

            except:
                print('Cannot execute function my_speak_sms')
        # --------------------------------------------------------------

        def my_sms_date():
            try:
                e_account_ssid_v=e_account_ssid.get()
                e_auth_v=e_auth.get()
                e_twilio_no_v=e_twilio_no.get()
                e_to_number_v=e_to_number.get()
                my_date=e_date.get()
                e_name_v=e_name.get()

                if len(my_date) >= 6:
                    conn=sqlite3.connect('attendance.db')
                    c=conn.cursor()
                    c.execute(""" SELECT fname,time,date from my_student WHERE date =?""",(my_date,))
                    today=c.fetchall()

                    counter=0


                    for data in range(0,len(today)):
                        # print(today[data][0],today[data][1],today[data][2])
                        counter=counter+1

                    my_body="Total Student Attended Lecuture on date {} is {}".format(my_date,counter)
                    # print(my_body)
                    # the following line needs your Twilio Account SID and Auth Token
                    client = Client(e_account_ssid_v,e_auth_v)
                    client.messages.create(to=e_to_number_v,
                                           from_= e_twilio_no_v,
                                           body=my_body)
                    p_sms=threading.Thread(target=my_speak_sms,args=('SMS sent to {} on  {}  Total Student attended {} lecture'.format(e_to_number_v,my_date,counter),))
                    p_sms.start()

                    messagebox.showinfo(self,"SMS send !")

                else:
                    messagebox.showinfo(self,"Please Enter Date !")
            except:
                print('cannot execute sms function ')

        def name_report_sms():
            try:
                e_account_ssid_v=e_account_ssid.get()
                e_auth_v=e_auth.get()
                e_twilio_no_v=e_twilio_no.get()
                e_to_number_v=e_to_number.get()
                my_date=e_date.get()
                e_name_v=e_name.get()
                print(e_name_v)
                if  len(e_name_v) >= 3:

                    # fetch data from the database and display that day !
                    conn=sqlite3.connect('attendance.db')
                    c=conn.cursor()
                    c.execute(""" SELECT fname,time,date from my_student WHERE fname =?""",(e_name_v,))
                    today=c.fetchall()

                    if len(today) == 0:
                        messagebox.showinfo(self,"No Record Found ! ")

                    else:
                        counter=0

                        for data in range(0,len(today)):

                            print(today[data][0],today[data][1],today[data][2])
                            counter=counter+1
                        my_body="Student{} Attended  {} Lecuture".format(e_name_v,counter)
                        print(my_body)
                        # the following line needs your Twilio Account SID and Auth Token
                        client = Client(e_account_ssid_v,e_auth_v)
                        client.messages.create(to=e_to_number_v,
                                               from_= e_twilio_no_v,
                                               body=my_body)
                    p_sms=threading.Thread(target=my_speak_sms,args=('SMS sent to {} and {} attended {} lecture'.format(e_to_number_v,e_name_v,counter),))
                    p_sms.start()
                    messagebox.showinfo(self,"SMS send !")
                else:
                    messagebox.showinfo(self,"Enter Name !")

            except:
                print('cannot execute function name_report_sms ')

        b_back = tk.Button(self, text="Back", bg="BlACK",fg="White",
                               command = lambda: controller.show_frame(PageThree))
        b_back.grid(row=6, column=3,sticky='NSEW', padx=10, pady=10)


class Cloud(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        # row 0--------------------------------------

        l_title=tk.Label(self, text="IoT Cloud Features ",
                         font="Arial,12")
        l_title.grid(row=0,column=0,columnspan=2, sticky="NSEW",padx=10,pady=10)

        clock = tk.Label(self, font=('times', 18, 'bold'),fg="RED")
        clock.grid(row=0,column=3, sticky="NSNESWSE",padx=8)

        def tick():
            time2=time.strftime('%H:%M:%S')
            clock.config(text=time2)
            clock.after(200,tick)
        tick()

        l_date=tk.Label(self,text='Enter Date')
        l_date.grid(row=1,column=0, sticky="NSNESWSE",padx=8)

        l_date_format=tk.Label(self,text='YYYY-MM-DD')
        l_date_format.grid(row=1,column=1, sticky="NSNESWSE",padx=14,pady=14)

        e_date=tk.Entry(self)
        e_date.grid(row=1,column=2, sticky="NSNESWSE",padx=14,pady=14)

        b_go=tk.Button(self,text='Go',bg='Black',fg='White',
                       command=lambda:cloud_date())
        b_go.grid(row=1,column=3, sticky="NSNESWSE",padx=14,pady=14)

        #========================================================
        l_name=tk.Label(self,text='Enter Name')
        l_name.grid(row=2,column=0, sticky="NSNESWSE",padx=8)

        e_name=tk.Entry(self)
        e_name.grid(row=2,column=2, sticky="NSNESWSE",padx=14,pady=14)

        b_go_name=tk.Button(self,text='Go',bg='Black',fg='White',
                            command=lambda :cloud_name())
        b_go_name.grid(row=2,column=3, sticky="NSNESWSE",padx=14,pady=14)

        b_back = tk.Button(self, text="Back", bg="BlACK",fg="White",
                           command = lambda: controller.show_frame(PageThree))
        b_back.grid(row=3, column=3,sticky='NSEW', padx=10, pady=10)


        def cloud_date():

            try:
                e_date_v = e_date.get()
                if len(e_date_v) == 0:
                    messagebox.showinfo(self,"Enter Date !")
                else:
                    c=threading.Thread(target=my_google,args=(e_date_v,))
                    c.start()

                    #my_google(e_date_v)
                    messagebox.showinfo(self,"Uploaded to Cloud!")
                    p_cloud=threading.Thread(target=my_speak_cloud,args=('Data was uploaded on google Spreadsheet for date {}'.format(e_date_v),))
                    p_cloud.start()
            except:
                messagebox.showinfo(self,"Error Cloud Server Crashed")


        def cloud_name():

            try:
                e_name_v = e_name.get()
                if len(e_name_v) == 0:
                    messagebox.showinfo(self,"Enter Name !")
                else:
                    m=threading.Thread(target=my_google_name,args=(e_name_v,))
                    m.start()
                    # my_google_name(e_name_v)

                    p_cloud=threading.Thread(target=my_speak_cloud,args=('Data was uploaded on google Spreadsheet ',))
                    p_cloud.start()
                    messagebox.showinfo(self,"Uploaded to Cloud!")
            except:
                p_cloud=threading.Thread(target=my_speak_cloud,args=('Server Crashed',))
                p_cloud.start()
                messagebox.showinfo(self,"Error Cloud Server Crashed")
                print('Cloud Crashed !')



        def my_speak_cloud(my_message):
            try:
                engine = pyttsx3.init()
                rate = engine.getProperty('rate')
                engine.setProperty('rate', rate)
                engine.say('{}'.format(my_message))
                engine.runAndWait()
                #rate = engine.getProperty('rate')
            except:
                print('Failed to exxceute my_speak_clud function ')








app = Attendance()
app.title('IoT Based RFID Attendance Monitoring Software ')
app.resizable(0, 0)
app.mainloop()


