#!/usr/bin/python
import datetime
import csv
import serial
import os
import sys

import tkinter as tk
import tkinter.filedialog as tkfd
import tkinter.scrolledtext as tkst
import tkinter.messagebox as tkmb
from tkinter import ttk

class App(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)

        self._job = None

        self.pack()
        self.master.title("Rate this sound")
        self.master.resizable(False, False)
        self.master.tk_setPalette(background='#ececec')
        x = round((self.master.winfo_screenwidth() - self.master.winfo_reqwidth()) / 2)
        y = round((self.master.winfo_screenheight() - self.master.winfo_reqheight()) / 3)
        self.master.geometry("+{}+{}".format(x, y))
        self.master.protocol('WM_DELETE_WINDOW', self.click_cancel)
        self.master.bind('<Return>', self.press_enter)
        self.master.bind('<Escape>', self.press_escape)

        self.serport = serial.Serial()


        dialog_frame = tk.Frame(self)
        dialog_frame.pack(padx=20, pady=15, anchor='w')
        tk.Label(dialog_frame, text='Save as:').grid(row=0, column=0, sticky='w')
        self.rating = tk.StringVar(value="?")
        self.working_path = tk.StringVar(value='')
        self.working_file = tk.StringVar(value='')
        self.working_file_entry = tk.Entry(dialog_frame, background='white', textvariable=self.working_file, width=24)
        self.working_file_entry.grid(row=0, column=1)
        tk.Button(dialog_frame, text='Choose file', command=self.file_picker).grid(row=0, column=2, sticky='w')
        tk.Label(dialog_frame, text='Serial port:').grid(row=1, column=0, sticky='w')
        self.serial_list = self.serial_ports()
        self.serial_list.insert(0, "Select a serial port")
        self.cbox = ttk.Combobox(dialog_frame, background='white', state='readonly', values=self.serial_list)
        self.cbox.set(self.serial_list[0])
        self.cbox.bind("<<ComboboxSelected>>", self.set_serport)
        self.cbox.grid(row=1, column=1)
        self.actionbutton_text = tk.StringVar(value="GO")
        self.actionbutton = tk.Button(dialog_frame, textvariable=self.actionbutton_text, width=40, state='disabled', command=self.click_action)
        self.actionbutton.grid(row=2, column=0, columnspan=3, padx=10, pady=(15,0), sticky='w')

        text_frame = tk.Frame(self)
        text_frame.pack()
        tk.Label(text_frame, text="Current rating:").pack()
        tk.Label(text_frame, textvariable=self.rating, font="Tahoma 72 bold").pack()

    def serial_ports(self):
        """ Lists serial port names

            :raises EnvironmentError:
                On unsupported or unknown platforms
            :returns:
                A list of the serial ports available on the system
        """
        if sys.platform.startswith('win'):
            ports = ['COM%s' % (i + 1) for i in range(256)]
        elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
            # this excludes your current terminal "/dev/tty"
            ports = glob.glob('/dev/tty[A-Za-z]*')
        elif sys.platform.startswith('darwin'):
            ports = glob.glob('/dev/tty.*')
        else:
            raise EnvironmentError('Unsupported platform')

        result = []
        for port in ports:
            try:
                s = serial.Serial(port = port, baudrate = 9600, timeout = 0.1)
                s.close()
                result.append(port)
            except (OSError, serial.SerialException):
                pass
        return result

    def read_serial(self):
        if self.serport.isOpen():
            self.serport.flush()
            self.rating.set(self.serport.readline().decode(sys.stdout.encoding).strip())
            self._job = self.master.after(10, self.read_serial)
            with open(self.working_path.get(), "a") as csv_file:
                csv_app = csv.writer(csv_file)
                if self.rating.get() != "":
                    csv_app.writerow([self.get_timestamp(), self.rating.get()])

        return 0

    def set_serport(self, event):
        portname = self.cbox.get()

        try:
            self.serport.close()
        except (OSError, serial.SerialException):
            pass

        self.cancel_reading()

        if "Select" not in portname :
            print(portname)
            self.serport = serial.Serial(port = portname, baudrate = 9600, timeout = 0.1)
            self.actionbutton.config(state='normal')
        else :
            print ("select a serial port")
            self.clear_serport()

    def clear_serport(self):
        try:
            self.serport.close()
        except (OSError, serial.SerialException):
            pass

        self.cancel_reading()

        self.rating.set("?")
        self.actionbutton_text.set("GO")
        self.actionbutton.config(state='disabled')

    def file_picker(self):
        self.working_path.set(tkfd.asksaveasfilename(parent=self))
        self.working_file.set(os.path.basename(self.working_path.get()))

    def click_action(self):
        if self.actionbutton_text.get() == "GO":
            if self.working_file.get() == "" :
                tkmb.showinfo("Output file",
                """Please select a file to write the data into.""")
            else :
                if self.serport.isOpen() == False :
                    tkmb.showinfo("Serial port",
                    """Please choose the appropriate serial port from the list.\n
                    You can disconnect and reconnect the slider to find out which
                    serial port is associated with it.""")
                else :
                    self.actionbutton_text.set("STOP")
                    self.read_serial()
        elif self.actionbutton_text.get() == "STOP" :
            self.actionbutton_text.set("GO")
            self.cancel_reading()
            # safe the data

    def cancel_reading(self):
        if self._job is not None:
            self.master.after_cancel(self._job)
            self._job = None

    def click_cancel(self):
        self.serport.close()
        self.master.destroy()

    def press_enter(self, event):
        self.actionbutton.invoke()

    def press_escape(self, event):
        self.click_cancel()

    def get_timestamp(self):
        firstpart = '{0:%H:%M:%S:}'.format(datetime.datetime.now())
        secondpart = '{}'.format(int(int('{0:%f}'.format(datetime.datetime.now())) / 100000))

        return firstpart + secondpart


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    app.mainloop()
