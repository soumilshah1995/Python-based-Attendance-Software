import binascii
import serial


def read_card(read_card,e_baud_v):                                         #  Read RFID card

    #print('''
    #        Place yoour card on reader
    #     ''')

    ser = serial.Serial(read_card,e_baud_v, timeout=2)
    data_raw = ser.readline()
    x = binascii.hexlify(data_raw)
    q = x.decode("ascii")  #converting scanned data
    #print('Remove Card')
    if len(q) > 9:
        data_key = q[4:27]
        return data_key