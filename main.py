import pyodbc
from tkinter import *
import time
from pymodbus.client import ModbusSerialClient as ModbusClient
import threading

root = Tk()

#title of the gui
root.title("Quenching Tank Temperature Monitoring System")

#Window size
#root.geometry("455x233")

#status bar for MOXA
stat =Label(root, relief="sunken",font=("Arial",12), anchor="w")
stat.grid(row="20", column="0")
moxaIp = Label(root,text="MOXA IP: 10.7.228.187")
moxaIp.grid(row="20",column="1")

#Label Widget for Heading
heading = Label(root, text="Quenching Tank Temperature Monitoring System", font=("Arial",24))
heading.grid(row=0, column=0, columnspan=2, sticky='ew')

#Label Widget for Quench Tanks
qt2 = Label(root, text="Quench Tank 2",font=("Arial",18)).grid(row=3, column=0)
qt3= Label(root, text="Quench Tank 3",font=("Arial",18)).grid(row=6, column=0)
qt4 = Label(root, text="Quench Tank 4",font=("Arial",18)).grid(row=9, column=0)
qt5 = Label(root, text="Quench Tank 5",font=("Arial",18)).grid(row=11, column=0)

#widgets for temperature values
temp2 = Label(root, relief="solid",font=("Arial",18))
temp3 = Label(root, relief="solid",font=("Arial",18))
temp4 = Label(root, relief="solid",font=("Arial",18))
temp5 = Label(root, relief="solid",font=("Arial",18))

#Temperature Labels
temp2.grid(row=3, column=1)
temp3.grid(row=6, column=1)
temp4.grid(row=9, column=1)
temp5.grid(row=11, column=1)

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
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    cursor.execute(f"INSERT INTO quenchTanksTemp(QT2,QT3,QT4,QT5,date_time) VALUES ({qT2Temp},{qT3Temp},{qT4Temp},{qt5Temp},GETDATE())") 
    conn.commit()
    conn.close()

def readTemperature(client,id):
    while True: 
        try:
            #print("Connecting to the server...")
            connection = client.connect()
            if(connection==True):
                stat.config(text=str("Successfully connected to server."))
            
                #Quench Tank 2 Temperature
                #inpReg2  = client.read_input_registers(0x06,1,unit=id)
                #qT2Temp = (inpReg2.registers[0]/10)
                qT2Temp=22.3 #remove this later
                #print("QT 2 Temperature:",qT2Temp,"Deg C")
                temp2.config(text=str("inactive") + " Deg C")
                

                #Quench Tank 3 Temperature
                inpReg3 = client.read_input_registers(0x06,1,unit=2)
                qT3Temp = (inpReg3.registers[0]/10)
                temp3.config(text=str(qT3Temp) + " Deg C")
                

                #Quench Tank 4 Temperature
                #inpReg4 = client.read_input_registers(0x06,3,unit=3)
                #qT4Temp = (inpReg4.registers[0]/10)
                qT4Temp = 44.4 #remove this later
                temp4.config(text=str("inactive") + " Deg C")
                

                #Quench Tank 5 Temperature
                #inpReg5 = client.read_input_registers(0x06,4,unit=4)
                #qT5Temp = (inpReg5.registers[0]/10)
                qt5Temp = 55.5 #remove this later
                temp5.config(text=str("inactive") + " Deg C")
                

                #dump to SQL
                insert_temperature_to_db(conn_str, qT2Temp,qT3Temp,qT4Temp,qt5Temp)
                
                #close the connection
                client.close()
                root.update()
                time.sleep(10)
            else:
                raise Exception("Unable to open serial port")
        except Exception as e:
            #raise this exception if the DP 9 Connecter is disconnected.(Failed to read the registers)

            print(e)
            stat.config(text=f"Disconnected: {e}")
            time.sleep(10)  # Wait before trying to reconnect
               
def main():
    client = ModbusClient(method = 'rtu', port='COM4', stopbits = 1, bytesize = 8, parity = 'N' , baudrate= 9600)
    threading.Thread(target=readTemperature, args=(client,1)).start()

if __name__ == "__main__":
    main()
    root.mainloop()