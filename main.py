#merging from main to frontend(tkinterFinal.py) and adding to new file merged
#Added the SQL error handling functionality
#Added and placed the graph buttons
#Graph plotting functionality yet to explored.

import pyodbc
from tkinter import *
import time
#import pymodbus
from pymodbus.pdu import ModbusRequest
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
from pymodbus.transaction import ModbusRtuFramer
import threading

#dump to SQL Database
#connection string
conn_str = (
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=10.7.228.186;" #Add host IP to which SQL Server is connected here
    "Port=1433;"  # Replace with your host machine's IP
    "Database=QuenchTank;"
    "UID=jack;"  # Replace with your SQL Server username
    "PWD=jack123;"  # Replace with your SQL Server password
)
def insert_temperature_to_db(conn_str,qT2Temp,qT3Temp,qT4Temp,qt5Temp):
    try:
        conn = pyodbc.connect(conn_str)
        database_connection_label.configure(text="Database Connection: CONNECTED", foreground='green')
        cursor = conn.cursor()

        # Replace None with 0 for all temperature variables
        #all the temp values will be 0 if no value is read.
        qT2Temp = 0 if qT2Temp is None else qT2Temp
        qT3Temp = 0 if qT3Temp is None else qT3Temp
        qT4Temp = 0 if qT4Temp is None else qT4Temp
        qt5Temp = 0 if qt5Temp is None else qt5Temp

        sql_query = f"INSERT INTO quenchTanksTemp(QT2,QT3,QT4,QT5,date_time) VALUES ({qT2Temp},{qT3Temp},{qT4Temp},{qt5Temp},GETDATE())"
        print(sql_query)
        cursor.execute(sql_query)
        conn.commit()
        conn.close()
    except Exception as e:
        print(e)
        database_connection_label.configure(text="Database Connection: DISCONNECTED", foreground='red')


root = Tk()
root.configure(background="black")
root.title("Quenching Tank Temperature Monitoring System")
root.geometry("1280x720")

#display label
heading_label = Label(root, text="QUENCHING TANK TEMPERATURE MONITORING SYSTEM", background="orange", font=('Arial','30','bold'), foreground="black")
QT2_label = Label(root, text='Quenching Tank 2', background="skyblue",font=('Arial','30','bold'), foreground="black")
QT3_label = Label(root, text='Quenching Tank 3', background="skyblue", font=('Arial','30','bold'), foreground="black")
QT4_label = Label(root, text='Quenching Tank 4', background="skyblue", font=('Arial','30','bold'), foreground="black")
QT5_label = Label(root, text='Quenching Tank 5', background="skyblue", font=('Arial','30','bold'), foreground="black")

#temperature values
QT2_temp_label = Label(root, background="blue", font=('Arial', '50','bold'), foreground="white")
QT3_temp_label = Label(root, background="blue", font=('Arial', '50','bold'), foreground="white")
QT4_temp_label = Label(root, background="blue", font=('Arial', '50','bold'), foreground="white")
QT5_temp_label = Label(root, background="blue", font=('Arial', '50','bold'), foreground="white")

#MODBUS PID configuration
QT2_Modbus_COM_port_label = Label(root, text="SLAVE ID: 1", font=('Arial','12','bold'), foreground="black")
QT3_Modbus_COM_port_label = Label(root, text="SLAVE ID: 2", font=('Arial','12','bold'), foreground="black")
QT4_Modbus_COM_port_label = Label(root, text="SLAVE ID: 3", font=('Arial','12','bold'), foreground="black")
QT5_Modbus_COM_port_label = Label(root, text="SLAVE ID: 4", font=('Arial','12','bold'), foreground="black")

#MODBUS Configuration
modbus_config_label = Label(root, text="METHOD = RTU / STOPBITS = 1 / DATA BITS = 8 / PARITY = NONE / BAUDRATE=9600", font=('Arial bold italic', '9'), foreground="white", background="black")

#MOXA Connetion Label
moxa_connection_label= Label(root, font=('Arial Bold Italic', '9'), foreground="white", background="black")

#database Connection Label
database_connection_label = Label(root, text="Database Connection: CONNECTING", font=('Arial bold italic', '9'), foreground="white", background="black")

#credits labels
credits_label = Label(root, text="Developed By: YASH PAWAR", font=('Arial Bold Italic','9'),foreground="white", background="black")

#Mentor label
mentor_label = Label(root, text="Guided By: SUNIL SINGH", font=('Arial Bold Italic','9'),foreground="white", background="black")


#grid definition

#configuring the number of colums
root.columnconfigure(0, weight = 1)
root.columnconfigure(1, weight = 2)
root.columnconfigure(2, weight = 1) 
root.columnconfigure(3, weight = 2)
root.columnconfigure(4, weight = 1)  

#configuring the number of rows
for i in range(14):
    root.rowconfigure({i},weight=1)

#WIDGET PLACEMENT 
#place heading widget
heading_label.grid(row=0, column=0, sticky="nsew", columnspan=8, pady=(0,30))

#place QT Label widget
QT2_label.grid(row=2,column=1, sticky="nsew")
QT3_label.grid(row=2,column=3, sticky="nsew")
QT4_label.grid(row=7,column=1, sticky="nsew")
QT5_label.grid(row=7,column=3, sticky="nsew")

#place temperature widget
QT2_temp_label.grid(row=3, column=1, sticky="nsew")
QT3_temp_label.grid(row=3, column=3, sticky="nsew")
QT4_temp_label.grid(row=8, column=1, sticky="nsew")
QT5_temp_label.grid(row=8, column=3, sticky="nsew")

#place Modbus PID configuration widget
QT2_Modbus_COM_port_label.grid(row=4, column=1, sticky="nsew")
QT3_Modbus_COM_port_label.grid(row=4, column=3, sticky="nsew")
QT4_Modbus_COM_port_label.grid(row=9, column=1, sticky="nsew")
QT5_Modbus_COM_port_label.grid(row=9, column=3, sticky="nsew")

#place database uplink label
database_connection_label.grid(row=11, column=1, columnspan=3, sticky="w")

#place MOXA Connection Label
moxa_connection_label.grid(row=12, column=1, sticky="w",columnspan=3)

#place the modbus config label
modbus_config_label.grid(row=13, column=1, sticky="w",columnspan=3)

#place the credits label
credits_label.grid(row=12, column=3, sticky ='e')

#place the mentor label
mentor_label.grid(row=13, column=3, sticky ='e')

#place the graph buttons
#graph buttons
frame1 = Frame(root)
frame1.grid(row=5, column=1, pady=10)
QT2_Graph = Button(frame1, text="QT2 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue', foreground='black').pack(expand=True)


frame2 = Frame(root)
frame2.grid(row=5, column=3, pady=10)
QT3_Graph = Button(frame2, text="QT3 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue').pack(expand=True)


frame3 = Frame(root)
frame3.grid(row=10, column=1, pady=10)
QT4_Graph = Button(frame3, text="QT4 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue').pack(expand=True)


frame4 = Frame(root)
frame4.grid(row=10, column=3, pady=10)
QT5_Graph = Button(frame4, text="QT5 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue').pack(expand=True)


def readTemperature(client):
    while True: 
        qT2Temp = qT3Temp = qT4Temp = qt5Temp = None  # Initialize the temperature variables
        try:
            #print("Connecting to the server...")
            connection = client.connect()
            if(connection==True):
                moxa_connection_label.config(text=str("MOXA: CONNECTED..! IP: 10.7.228.186"), foreground="Green")
            
                #Quench Tank 2 Temperature
                #Quench Tank 3 Temperature
                try:
                    inpReg2 = client.read_input_registers(0x06,1,unit=2)
                    qT2Temp = (inpReg2.registers[0]/10)
                    QT2_temp_label.config(text=str(qT2Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT2Temp=0
                    QT2_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                

                #Quench Tank 3 Temperature
                try:
                    inpReg3 = client.read_input_registers(0x06,1,unit=3)
                    qT3Temp = (inpReg3.registers[0]/10)
                    QT3_temp_label.config(text=str(qT3Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT3Temp=0
                    QT3_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                
                #Quench Tank 4 Temperature
                try:
                    inpReg4 = client.read_input_registers(0x06,1,unit=4)
                    qT4Temp = (inpReg4.registers[0]/10)
                    QT4_temp_label.config(text=str(qT4Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT4Temp=0
                    QT4_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                

                #Quench Tank 5 Temperature
                try:
                    inpReg5 = client.read_input_registers(0x06,1,unit=5)
                    qT5Temp = (inpReg5.registers[0]/10)
                    QT5_temp_label.config(text=str(qT5Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT5Temp=0
                    QT5_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                

                #dump to SQL
                insert_temperature_to_db(conn_str, qT2Temp,qT3Temp,qT4Temp,qt5Temp)
                
                #close the connection
                client.close()
                root.update()
                time.sleep(2)
                insert_temperature_to_db(conn_str, qT2Temp,qT3Temp,qT4Temp,qt5Temp)
            else:
                QT2_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                QT3_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                QT4_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                QT5_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                raise Exception(moxa_connection_label.config(text=str("MOXA: DISCONNECTED..! IP: 10.7.228.186"),foreground="red"))
        except Exception as e:
            #raise this exception if the DP 9 Connecter is disconnected from MOXA.(Failed to read the registers)
            #print(e)
            time.sleep(2)  # Wait before trying to reconnect

#main loop of the program
def main():
    client = ModbusClient(method = 'rtu', port='COM4', stopbits = 1, bytesize = 8, parity = 'N' , baudrate= 9600)
    threading.Thread(target=readTemperature, args=(client,)).start()

if __name__ == "__main__":
    main()
    root.mainloop()