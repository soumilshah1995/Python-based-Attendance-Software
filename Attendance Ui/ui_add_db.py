import sqlite3
import serial

def add_data_database(e_name_v, e_name_m_v,e_name_l_v,
                      e_email_v, e_id_v, e_phone_v,
                      e_rfid_v, e_gender_v):

    conn = sqlite3.connect('UI_User.db')
    c=conn.cursor()

    c.execute(""" 
    CREATE TABLE IF NOT EXISTS my_student 
    (fname TEXT,mname TEXT,lname TEXT,email TEXT,id TEXT,phone TEXT,rfid TEXT,gender TEXT)""")



    c.execute("""INSERT INTO my_student 
    (fname, mname, lname, email, id, phone, rfid, gender) 
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)""", (e_name_v, e_name_m_v,e_name_l_v,e_email_v, e_id_v, e_phone_v,e_rfid_v, e_gender_v,))


    conn.commit()
    c.close()
    conn.close()


