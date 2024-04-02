'''
    Email trigger for MOXA disconnection

'''
from turtle import back
from pandas import Timestamp
import pyodbc
from tkinter import *
import time
#import pymodbus
from pymodbus.pdu import ModbusRequest
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
from pymodbus.transaction import ModbusRtuFramer
import threading

#graph plotting dependencies
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import(
    FigureCanvasTkAgg, NavigationToolbar2Tk)
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from matplotlib.dates import DateFormatter

from datetime import timedelta

#for reading saved COM port
import json
from tkinter import messagebox

#for accessing the system available COM ports
import serial.tools.list_ports

#For temp queue
from queue import Queue

#email triggering dependencies
from email.message import EmailMessage
from multiprocessing import context
from password import mypassword
import ssl
import smtplib

email_sender = "practicemail410@gmail.com"
email_password = mypassword
email_receiver = "pyash632002@gmail.com"
email_sent = False #email flag

#GLOBAL VARIABLES 
global settings_password
settings_password  = "mypassword"

global tempQueue
tempQueue = Queue()

qT2Temp = None
qT3Temp = None
qT4Temp = None 
qT5Temp = None  

#dump to SQL Database
#connection string
conn_str = (
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=10.7.228.186;" #Add host IP to which SQL Server is connected here
    "Port=1433;"  # add the port no. to which sql server listens
    "Database=QuenchTank;"  #database name
    "UID=jack;"  # Replace with your SQL Server username
    "PWD=jack123;"  # Replace with your SQL Server password
)
    
def insert_temperature_to_db(conn_str,qT2Temp,qT3Temp,qT4Temp,qt5Temp):
    try:
        conn = pyodbc.connect(conn_str, timeout=5)
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

def get_saved_com_port():
    try: 
        with open('settings.json','r') as json_file:
            settings = json.load(json_file)
            #print(settings)
            return settings.get('ComPort','COM1')
        
    except FileNotFoundError:
            messagebox.showinfo("Settings not found", "Please select the COM port manually.")
            return ""

def get_available_comports():
    ports = serial.tools.list_ports.comports()
    available_ports = [port.device for port in ports]
    return available_ports

def create_settings_window():
    settings_window = Toplevel(root)
    settings_window.configure(background="black")
    settings_window.title("Settings")
    settings_window.resizable(False, False)
    settings_window.geometry("600x300")

    settings_window.columnconfigure(0, weight=1)
    settings_window.columnconfigure(1, weight=3)

    for i in range(7):
        settings_window.rowconfigure({i})
        
    #Settings window heading
    settings_title_label = Label(settings_window, text="Settings", background="orange",font=('Arial','30','bold'), foreground="black")
    settings_title_label.grid(row=0, column=0, sticky='nsew', columnspan=2)

    comport_set_label = Label(settings_window, text="COM PORT: ", background="black",font=('Arial','15','bold'), foreground="white")
    comport_set_label.grid(row=2, column=0, sticky="w", pady=20, padx=20)

    settings_password_label = Label(settings_window, text="Password: ",font=('Arial','15','bold'), background="black", foreground="white")
    settings_password_label.grid(row=4, column=0, sticky="w", pady=20, padx=20)

    settings_password_input = Entry(settings_window, width=11, show="*")
    settings_password_input.grid(row=4, column=0, sticky="e")
    
    #Invalid password label
    
    invalid_password_label = Label(settings_window, text="", background="black")
    invalid_password_label.grid(row=6, column=0, sticky="w", padx=20)

    #Read the available COM Ports from system
    com_drop_options_list = get_available_comports()
    
    com_drop_clicked = StringVar()

    settings_apply_button = Button(settings_window, text="Apply", font=('Arial','12','bold'), background='light blue', foreground='black', command=lambda: apply_settings(settings_password_input.get(), invalid_password_label, com_drop_clicked))
    settings_apply_button.grid(row=6, column=1, sticky="w", padx=20)
    #selected_com_port = com_drop_clicked.get() #access the selected com port value using this

    com_drop_clicked.set(com_drop_options_list[0])
    com_dropdown = OptionMenu(settings_window, com_drop_clicked, *com_drop_options_list)

    com_dropdown.configure(background='light blue', highlightthickness=0)
    com_dropdown.grid(row=2, column=0, sticky="e")

def apply_settings(password, invalid_password_label, com_drop_clicked):
        if password != settings_password:
            invalid_password_label.config(text="Invalid Password!", font=('Arial','12','bold'), foreground="red", background="white")
            return
        else:
            selected_com_port = com_drop_clicked.get()
            invalid_password_label.config(text="Settings Applied", font=('Arial','12','bold'), foreground="green", background="white")
            with open('settings.json', 'w') as json_file:
                json.dump({'ComPort': selected_com_port}, json_file)
                update_modbus_config_label(selected_com_port)
                #print(selected_com_port)
            
def open_graph_window(tempVal, graph_name):
    plt.style.use('dark_background')
    graph_window = Toplevel(root)
    graph_window.resizable(False, False)
    graph_window.title(graph_name)
    graph_window.iconbitmap('Images\QTTMS logo.ico')
    #Initialize Tkinter and Matplotlib Figure
    fig_graph, axis_graph = plt.subplots()

    #the x and y values will be stored in the following
    # global x_vals, y_vals
    x_vals = []
    y_vals = []
    
    def animate(i):
        try:
            conn = pyodbc.connect(conn_str, timeout=5)
            cursor = conn.cursor()
            cursor.execute("SELECT CURRENT_TIMESTAMP")
        except Exception as e:
            messagebox.showinfo("Database not connected."," Unable to fetch server date and time")
        row = cursor.fetchone()

        if row and tempVal is not None:
            x_vals.append(row[0])
            y_vals.append(tempVal)

        #clear axis
        plt.cla()
        axis_graph.plot(x_vals, y_vals, color='orange', linewidth=2)

        if len(x_vals) > 10:
            axis_graph.set_xlim(left=x_vals[-10], right=x_vals[-1])
        axis_graph.set_title('REAL TIME TREND')

        if y_vals:
            axis_graph.set_ylim([0, max(y_vals) + 10])
        else:
            axis_graph.set_xlim([0, 10])

        axis_graph.set_xlabel('DATE & TIME')
        axis_graph.set_ylabel('TEMP IN °C')
        #axis_graph.plot(x_vals,y_vals)

        date_format = DateFormatter('%Y-%m-%d %H:%M')
        axis_graph.xaxis.set_major_formatter(date_format)
        
        for label in axis_graph.get_xticklabels():
            label.set_rotation(45)
            label.set_fontsize(8)

        # Add padding to the y-axis
        plt.margins(y=0.2) 
        conn.commit()
        conn.close() 

    #call to animate function
    ani = FuncAnimation(plt.gcf(), animate, interval=1000)

    #Tkinter Application Window
    frame = Frame(graph_window)
    frame.grid(sticky='ew')  # Use grid for frame

    label = Label(frame, text = graph_name, font=('Arial','32','bold'), background="orange", foreground="black")  # Add label to frame
    label.pack(fill='both', expand=True)  # Use grid for label

    #create canvas
    canvas = FigureCanvasTkAgg(fig_graph, master=graph_window)
    canvas.get_tk_widget().grid(sticky='nsew')  # Use grid for canvas

    # Increase bottom padding
    plt.subplots_adjust(bottom=0.3)

    #Create Toolbar
    # Create custom toolbar
    class CustomToolbar(NavigationToolbar2Tk):
        def set_message(self, s):
            s = s.replace('x', 'Date & Time').replace('y', 'Temp in °C')
            super().set_message(s)
    # Create toolbar frame 
    toolbar_frame = Frame(graph_window)
    toolbar_frame.grid()  # Use grid for toolbar_frame
    
    # Create custom toolbar
    toolbar = CustomToolbar(canvas, toolbar_frame)
    toolbar.update()
    toolbar.grid(sticky='w')
     
    canvas.draw()

    def close_graph_window():
        plt.close(fig_graph)
        graph_window.destroy()
    graph_window.protocol("WM_DELETE_WINDOW", close_graph_window)

    graph_window.mainloop()

def readTemperature(modbus_client):
    global qT2Temp
    global qT3Temp
    global qT4Temp
    global qT5Temp
    qT2Temp = qT3Temp = qT4Temp = qT5Temp = None
    global tempQueue
    global email_sent
    while True: 
        try:
            #print("Connecting to the server...")
            connection = modbus_client.connect()
            if(connection==True):
                moxa_connection_label.config(text=str("MOXA: CONNECTED..! IP: 10.7.228.186"), foreground="Green")
                global email_sent
                email_sent = False
                #Quench Tank 2 Temperature
                try:
                    inpReg2 = modbus_client.read_input_registers(0x06,1,unit=2)
                    qT2Temp = (inpReg2.registers[0]/10)
                    QT2_temp_label.config(text=str(qT2Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT2Temp=0
                    QT2_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                
                #Quench Tank 3 Temperature
                try:
                    inpReg3 = modbus_client.read_input_registers(0x06,1,unit=3)
                    qT3Temp = (inpReg3.registers[0]/10)
                    QT3_temp_label.config(text=str(qT3Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT3Temp=0
                    QT3_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                
                #Quench Tank 4 Temperature
                try:
                    inpReg4 = modbus_client.read_input_registers(0x06,1,unit=4)
                    qT4Temp = (inpReg4.registers[0]/10)
                    QT4_temp_label.config(text=str(qT4Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT4Temp=0
                    QT4_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))
                
                #Quench Tank 5 Temperature
                try:
                    inpReg5 = modbus_client.read_input_registers(0x06,1,unit=5)
                    qT5Temp = (inpReg5.registers[0]/10)
                    QT5_temp_label.config(text=str(qT5Temp)+"°C",background="blue",font=('Arial','50','bold'))
                except Exception as e:
                    qT5Temp=0
                    QT5_temp_label.config(text="PID Disconnected", background="red", font=('Arial','20','bold'))

                tempQueue.put((qT2Temp, qT3Temp, qT4Temp, qT5Temp))
                #close the modbus connection
                modbus_client.close()
                root.update()
                time.sleep(2)
                
            else:
                QT2_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                QT3_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                QT4_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                QT5_temp_label.config(text="Moxa Disconnected", background="red", font=('Arial','20','bold'))
                if not email_sent:
                    try:
                        conn = pyodbc.connect(conn_str, timeout=5)
                        cursor = conn.cursor()
                        cursor.execute("SELECT CURRENT_TIMESTAMP")
                        row = cursor.fetchone()

                        timestamp = row[0].strftime("%Y-%m-%d %H:%M:%S")
                    except Exception as e:
                        print(f"An error occurred: {e}")
                        timestamp = "unknown"

                    timestamp = row[0].strftime("%Y-%m-%d %H:%M:%S")
                    email_subject = "QTTMS Alerts - MOXA Server Disconnected from the system!"
                    email_body =  f'''Attention: The MOXA converter with  IP 10.7.228.187 was disconnected from the system at {timestamp} \n Location - Gate No. 10, L&T Special Steels and Heavy Forgings, Hazira, Surat.
                    '''
                    email_message = EmailMessage()
                    email_message['From'] = email_sender
                    email_message['To'] = email_receiver
                    email_message['subject'] = email_subject
                    email_message.set_content(email_body)

                    context = ssl.create_default_context()

                    with smtplib.SMTP_SSL('smtp.gmail.com',465, context=context) as smtp:

                        smtp.login(email_sender, email_password)
                        smtp.sendmail(email_sender, email_receiver, email_message.as_string())
                    email_sent = True  # set the flag to True after the email is sent
                    raise Exception(moxa_connection_label.config(text=str("MOXA: DISCONNECTED..! IP: 10.7.228.186"),foreground="red"))
        except Exception as e:
            #raise this exception if the DP 9 Connecter is disconnected from MOXA.(Failed to read the registers)
            #print(e)
            time.sleep(2)  # Wait before trying to reconnect

def dump_to_db():
    # Initialize the temperature variables
    global tempQueue
    while True:
        qT2Temp, qT3Temp, qT4Temp, qT5Temp = tempQueue.get()
        print(f"{qT2Temp},{qT3Temp},{qT4Temp},{qT5Temp}")
        insert_temperature_to_db(conn_str, qT2Temp,qT3Temp,qT4Temp,qT5Temp)
        time.sleep(5)

root = Tk()
root.configure(background="black")
root.title("Quenching Tank Temperature Monitoring")
root.geometry("1280x720")
root.iconbitmap('Images\QTTMS logo.ico')

#display label
heading_label = Label(root, text="QUENCHING TANK TEMPERATURE MONITORING", background="orange", font=('Arial','30','bold'), foreground="black")
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
def update_modbus_config_label(com_port):
    modbus_config_label = Label(root, text=f"METHOD = RTU / STOPBITS = 1 / DATA BITS = 8 / PARITY = NONE / BAUDRATE=9600 /{com_port}", font=('Arial bold italic', '9'), foreground="white", background="black")
    modbus_config_label.grid(row=13, column=1, sticky="w",columnspan=3)

# Initially create the label without com_port
modbus_config_label = Label(root,text=f"METHOD = RTU / STOPBITS = 1 / DATA BITS = 8 / PARITY = NONE / BAUDRATE=9600 /", font=('Arial bold italic', '9'), foreground="white", background="black")
modbus_config_label.grid(row=13, column=1, sticky="w",columnspan=3)

#MOXA Connetion Label
moxa_connection_label= Label(root, font=('Arial Bold Italic', '9'), foreground="white", background="black")

#database Connection Label
database_connection_label = Label(root, text="Database Connection: CONNECTING", font=('Arial bold italic', '9'), foreground="white", background="black")

#credits labels
credits_label = Label(root, text="Developed By: YASH PAWAR", font=('Arial Bold Italic','9'),foreground="white", background="black")

#Mentor label
mentor_label = Label(root, text="Guided By: SUNIL SINGH", font=('Arial Bold Italic','9'),foreground="white", background="black")

#grid definition

#configuring the number of columns
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

#place the credits label
credits_label.grid(row=12, column=3, sticky ='e')

#place the mentor label
mentor_label.grid(row=13, column=3, sticky ='e')

#place the graph buttons
#graph buttons
frame1 = Frame(root)
frame1.grid(row=5, column=1, pady=10)
QT2_Graph = Button(frame1, text="QT2 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue', foreground='black', command=lambda: open_graph_window(tempVal=qT2Temp,graph_name="QT2 GRAPH")).grid(row=0, column=0)

frame2 = Frame(root)
frame2.grid(row=5, column=3, pady=10)
QT3_Graph = Button(frame2, text="QT3 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue', foreground='black', command=lambda: open_graph_window(tempVal=qT3Temp, graph_name="QT3 GRAPH")).grid(row=0, column=0)

frame3 = Frame(root)
frame3.grid(row=10, column=1, pady=10)
QT4_Graph = Button(frame3, text="QT4 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue', foreground='black', command=lambda: open_graph_window(tempVal=qT4Temp, graph_name="QT4 GRAPH")).grid(row=0, column=0)

frame4 = Frame(root)
frame4 = Frame(root)
frame4.grid(row=10, column=3, pady=10)
QT5_Graph = Button(frame4, text="QT5 GRAPH", width=20, height=2, font=('Arial','12','bold'), background='light blue', foreground='black', command=lambda: open_graph_window(tempVal=qT5Temp, graph_name="QT5 GRAPH")).grid(row=0, column=0)

#add settings button
settings_frame = Frame(root)
settings_frame.grid(row=11, column=2)
settings_button = Button(settings_frame, text="SETTINGS", width=10, height=2, font=('Arial','12','bold'), background='light blue', foreground='black', command=lambda: create_settings_window()).pack(expand=True)
 
#main loop of the program
def main():
    
    com_port = get_saved_com_port()
    update_modbus_config_label(com_port)

    modbus_client = ModbusClient(method = 'rtu', port=com_port, stopbits = 1, bytesize = 8, parity = 'N' , baudrate= 9600)
    read_temp_thread = threading.Thread(target=readTemperature, args=(modbus_client,))
    read_temp_thread.start()

    dump_to_db_thread = threading.Thread(target=dump_to_db)
    dump_to_db_thread.start()

if __name__ == "__main__":
    main()
    root.mainloop()