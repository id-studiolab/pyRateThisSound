#!/usr/bin/python
import threading

import wave
import pyaudio
import time
import csv
import serial
import os
import sys

import tkinter as tk
import tkinter.filedialog as tkfd
import tkinter.scrolledtext as tkst
import tkinter.messagebox as tkmb
from tkinter import ttk

class AudioThread(threading.Thread):
    """ Ask the thread to stop by calling its join() method. """
    def __init__(self, audio_path):
        super(AudioThread, self).__init__()
        self.stoprequest = threading.Event()
        self.audio_path = audio_path

    def run(self):
        """"
        As long as we weren't asked to stop, try to take new tasks from the
        queue. The tasks are taken with a blocking 'get', so no CPU
        cycles are wasted while waiting.
        Also, 'get' is given a timeout, so stoprequest is always checked,
        even if there's nothing in the queue.
        """
        f = wave.open(self.audio_path, "rb")
        p = pyaudio.PyAudio()

        stream = p.open(format = p.get_format_from_width(f.getsampwidth()),\
                        channels = f.getnchannels(),\
                        rate = f.getframerate(),\
                        output = True)

        data = f.readframes(1024)

        while data and not self.stoprequest.isSet():
            stream.write(data)
            data = f.readframes(1024)

        stream.stop_stream()
        stream.close()
        p.terminate()

    def join(self, timeout=None):
        self.stoprequest.set()
        super(AudioThread, self).join(timeout)


class App(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)

        self._check_audio_job = None
        self._read_serial_job = None
        self.rating_list = []
        self.serport = serial.Serial()

        self.pack()
        self.master.title("Rate this sound v2")
        self.master.resizable(False, False)
        self.master.tk_setPalette(background='#ececec')
        x = round((self.master.winfo_screenwidth() - self.master.winfo_reqwidth()) / 2)
        y = round((self.master.winfo_screenheight() - self.master.winfo_reqheight()) / 3)
        self.master.geometry("+{}+{}".format(x, y))
        self.master.protocol('WM_DELETE_WINDOW', self.click_cancel)
        self.master.bind('<Return>', self.press_enter)
        self.master.bind('<Escape>', self.press_escape)

        dialog_frame = tk.Frame(self)
        dialog_frame.pack(padx=20, pady=15, anchor='w')
        tk.Label(dialog_frame, text='Serial port:').grid(row=0, column=0, sticky='w')
        self.serial_list = self.get_port_list()
        self.serial_list.insert(0, "Select a serial port")
        self.cbox = ttk.Combobox(dialog_frame, background='white', state='readonly', values=self.serial_list)
        self.cbox.set(self.serial_list[0])
        self.cbox.bind("<<ComboboxSelected>>", self.set_serport)
        self.cbox.grid(row=0, column=1)
        tk.Label(dialog_frame, text='Save as:').grid(row=1, column=0, sticky='w')
        self.rating = tk.StringVar(value="?")
        self.working_path = tk.StringVar(value='')
        self.working_file = tk.StringVar(value='')
        self.audio_path = tk.StringVar(value='')
        self.audio_file = tk.StringVar(value='')
        self.working_file_entry = tk.Entry(dialog_frame, background='white', textvariable=self.working_file, width=24)
        self.working_file_entry.grid(row=1, column=1)
        tk.Button(dialog_frame, text='Choose file', command=self.file_picker).grid(row=1, column=2, sticky='w')
        tk.Label(dialog_frame, text='Audio file:').grid(row=2, column=0, sticky='w')
        self.audio_file_entry = tk.Entry(dialog_frame, background='white', textvariable=self.audio_file, width=24)
        self.audio_file_entry.grid(row=2, column=1)
        tk.Button(dialog_frame, text='Choose file', command=self.audio_picker).grid(row=2, column=2, sticky='w')
        self.actionbutton_text = tk.StringVar(value="GO")
        self.actionbutton = tk.Button(dialog_frame, textvariable=self.actionbutton_text, width=40, state='disabled', command=self.click_action)
        self.actionbutton.grid(row=3, column=0, columnspan=3, padx=10, pady=(15,0), sticky='w')
        text_frame = tk.Frame(self)
        text_frame.pack()
        tk.Label(text_frame, text="Current rating:").pack()
        tk.Label(text_frame, textvariable=self.rating, font="Tahoma 72 bold").pack()
        tk.Label(text_frame, text="support: r.bekking@tudelft.nl", anchor='e').pack()

    def get_port_list(self):
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
            self._read_serial_job = self.master.after(10, self.read_serial)

            if self.rating.get() != "":
                self.rating_list.append([self.get_timestamp(), self.rating.get()])

        return 0

    def cancel_read_serial(self):
        if self._read_serial_job is not None:
            self.rating.set("?")
            self.master.after_cancel(self._read_serial_job)
            self._read_serial_job = None

    def set_serport(self, event):
        portname = self.cbox.get()

        try:
            self.serport.close()
        except (OSError, serial.SerialException):
            pass

        self.cancel_read_serial()

        if "Select" not in portname :
            self.serport = serial.Serial(port = portname, baudrate = 9600, timeout = 0.1)
            self.actionbutton.config(state='normal')
        else :
            self.clear_serport()

    def clear_serport(self):
        try:
            self.serport.close()
        except (OSError, serial.SerialException):
            pass

        self.cancel_reading()

        self.actionbutton_text.set("GO")
        self.actionbutton.config(state='disabled')

    def start_audio_playback(self):
        if self.audio_path.get():
            try:
                self.wt = AudioThread(self.audio_path.get())
                self.wt.start()
                self.check_audio()
            except Exception as e:
                print(e)
                tkmb.showwarning("Audio",
                """Cannot play audio file.\n
                Make sure that the file you're trying to play is a uncompressed
                wave file.""")

    def check_audio(self):
        if self.audio_file.get() and self.wt.isAlive():
            self._check_audio_job = self.master.after(10, self.check_audio)
        else:
            self.actionbutton.invoke()

        return

    def cancel_audio_playback(self):
        if self._check_audio_job is not None:
            self.wt.join() # Stop audio
            self.master.after_cancel(self._check_audio_job)
            self._check_audio_job = None

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
                    # Play audio
                    self.rating_list.clear()
                    self.actionbutton_text.set("STOP")
                    self.start_time = time.time()
                    self.start_audio_playback()
                    self.read_serial()
        elif self.actionbutton_text.get() == "STOP" :
            self.actionbutton_text.set("GO")
            self.cancel_audio_playback()
            self.cancel_read_serial()
            self.safe_to_file(self.working_path.get(), self.rating_list)

    def safe_to_file(self, filename, data):
        if filename:
            with open(filename, "a") as csv_file:
                csv_app = csv.writer(csv_file)
                for row in data:
                    csv_app.writerow(row)

        return

    def click_cancel(self):
        self.safe_to_file(self.working_path.get(), self.rating_list)
        self.serport.close()
        self.master.destroy()

    def get_timestamp(self):
        current_time = time.time()
        hours, rem = divmod(current_time - self.start_time, 3600)
        minutes, seconds = divmod(rem, 60)
        return ("{:0>2}:{:0>2}:{:05.2f}".format(int(hours), int(minutes), seconds))

    def press_enter(self, event):
        self.actionbutton.invoke()

    def press_escape(self, event):
        self.click_cancel()

    def file_picker(self):
        self.working_path.set(tkfd.asksaveasfilename(parent=self))
        self.working_file.set(os.path.basename(self.working_path.get()))

    def audio_picker(self):
        self.audio_path.set(
            tkfd.askopenfilename(
            parent=self,
            filetypes=(("Uncrompressed WAV files", "*.wav"),
                       ("All files", "*.*"))
            )
        )
        self.audio_file.set(os.path.basename(self.audio_path.get()))

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    app.mainloop()
