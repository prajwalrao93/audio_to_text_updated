from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox, filedialog
from tkinter import font
from ttkthemes import ThemedTk
from os import listdir
import os, openpyxl, sys, time, threading, shutil
import speech_recognition as sr
from ffmpy import FFmpeg
from tkinter.scrolledtext import ScrolledText


class helpScreen:
    def __init__(self, root):
        self.root = root
        self.root.minsize(500,500)
        self.root.resizable(False, False)
        self.root['bg'] = "#300a24"
        self.style = Style()
        self.style.theme_use("radiance")
        self.style.configure('TButton', background="#300a24")
        self.style.configure('TFrame', background="#300a24")
        self.style.configure('TLabel', foreground='white', background="#300a24")
        self.style.configure('TEntry', background="#300a24")
        
          
        self.font_def = font.Font(family="Helvetica", size=12, weight='bold', slant='italic')

        text = """
Method to use this app

1: Place the media files in the Media folder.
2: Run the Converter.bat file in the Media folder and wait for it to get completed.
3: Once the bat file has run, now run the speech_to_text.exe file.
4: Copy the folder path of the media folder and paste in the first open space
5: Copy the raw data file path along with file name in the second open space
6: Click on Submit button. Wait for the application to run and complete
7: Use Clear button to clear the fields
8: Use Exit button to exit the application
9: Use Help button for help on how to use the application

For more information contact
"""

        self.frame = Frame(self.root, style="TFrame")

        self.text1 = ScrolledText(self.frame, wrap=WORD, width=30, height=20)
        self.text1.pack(expand=True, side=TOP, padx=10, pady=10)
        self.text1.insert(INSERT, text)

        self.button = Button(self.frame, text="Quit", width=10, command=self.root.destroy, style="TButton")
        self.button.pack(expand=True, side=TOP)

        self.frame.pack(expand=True)
        


class MyApp():
    def __init__(self, root):
        self.root = root
        self.style = Style()
        self.style.theme_use("radiance")
        self.style.configure('TButton', background="#300a24")
        self.style.configure('TFrame', background="#300a24")
        self.style.configure('TLabel', foreground='white',background="#300a24")
        self.style.configure('TEntry', background="#300a24")
        
        self.root.minsize(500,500)
        self.root.resizable(False, False)
        self.root['bg'] = "#300a24"
        
        self.font_def = font.Font(family="Helvetica", size=12, weight='bold', slant='italic')

        self.frame1 = Frame(root, style="TFrame")
        self.frame2 = Frame(root, style="TFrame")
        self.frame3 = Frame(root, style="TFrame")
        self.frame4 = Frame(root, style="TFrame")
        self.frame5 = Frame(root, style="TFrame")
        self.frame6 = Frame(root, style="TFrame")

        self.label1 = Label(self.frame2, text="Media Folder Path: ", style="TLabel", font=self.font_def)
        self.label1.pack(expand=True, side=LEFT, padx=50)

        self.entry1 = Entry(self.frame2, width=40, style="TEntry")
        self.entry1.pack(expand=True, side=LEFT, padx=50)

        self.button5 = Button(self.frame2, text="...", style="TButton", width=2, command=self.select_directory)
        self.button5.pack(expand=True, side=LEFT, padx=50)

        self.label2 = Label(self.frame3, text="Raw Data File Path: ", style="TLabel", font=self.font_def)
        self.label2.pack(expand=True, side=LEFT, padx=50)

        self.entry2 = Entry(self.frame3, width=40, style="TEntry")
        self.entry2.pack(expand=True, side=LEFT, padx=50)

        self.button6 = Button(self.frame3, text="...", style="TButton", width=2, command=self.select_file)
        self.button6.pack(expand=True, side=LEFT, padx=50)

        self.button1 = Button(self.frame4, text="Submit", style="TButton", width=10, command=self.thread_func)
        self.button1.pack(expand=True, side=LEFT, padx=20, pady=10)

        self.button2 = Button(self.frame4, text="Clear", style="TButton", width=10, command=self.clearAll)
        self.button2.pack(expand=True, side=LEFT, padx=20, pady=10)

        self.button3 = Button(self.frame4, text="Exit", style="TButton", width=10, command=self.root.destroy)
        self.button3.pack(expand=True, side=LEFT, padx=20, pady=10)

        self.pbar = Progressbar(self.frame5, maximum=100, orient=HORIZONTAL, length = 350, mode='determinate')
        self.pbar['value'] = 0
        self.pbar.pack(expand=True)

        self.name = StringVar()
        self.name.set("")


        self.label3 = Label(self.frame6, textvariable=self.name, style="TLabel", font=self.font_def)
        self.label3.pack(expand=True, side=TOP)

        self.button4 = Button(self.frame6, text="Help", width=10, style="TButton", command=self.second_window)
        self.button4.pack(expand=True, side=TOP)


        self.frame1.pack(expand=True)
        self.frame2.pack(expand=True)
        self.frame3.pack(expand=True)
        self.frame4.pack(expand=True)
        self.frame5.pack(expand=True)
        self.frame6.pack(expand=True)

    def select_directory(self):
        self.wav_path = filedialog.askdirectory()
        self.entry1.delete(0, END)
        self.entry1.insert(0, self.wav_path)
        self.entry1.update

    def select_file(self):
        self.excel_path = filedialog.askopenfilename(filetypes=(("excel files", ".xlsx"), ("excel files", ".xlsm"),("all files",".")))
        self.entry2.delete(0, END)
        self.entry2.insert(0, self.excel_path)
        self.entry2.update

    def second_window(self):
        self.root1 = ThemedTk(theme="radiance")
        self.helpWindow = helpScreen(self.root1)
        self.root1.mainloop()


    def clearAll(self):
        self.entry1.delete(0, last=len(self.entry1.get()))
        self.entry1.update
        self.entry2.delete(0, last=len(self.entry2.get()))
        self.entry2.update

    
    def combined_text(self, updated_file):
        n = len(updated_file)
        r = sr.Recognizer()
        wit_ai_key = "WWEHIPV5SE6EXEAV6YE4QAJZE5LF22JW"
        if n == 1:
            with sr.AudioFile(os.path.join(self.mp3_path, updated_file[0])) as source:
                r.adjust_for_ambient_noise(source)
                audio_text = r.listen(source)
            try:
                text = r.recognize_wit(audio_text, key=wit_ai_key)
            except sr.UnknownValueError:
                text = "Audio not Clear"
                #shutil.copyfile(os.path.join(self.mp3_path, v), os.path.join(self.mp3_path, "NotRecognized", v))
            except sr.RequestError as e:
                text = "Audio not Recognized"
                #shutil.copyfile(os.path.join(self.mp3_path, v), os.path.join(self.mp3_path, "NotRecognized", v))
        else:
            converted_list = []
            for i in updated_file:
                time.sleep(1)
                with sr.AudioFile(os.path.join(self.mp3_path, i)) as source:
                    r.adjust_for_ambient_noise(source)
                    audio_text = r.listen(source)
                try:
                    ind_text = r.recognize_wit(audio_text, key=wit_ai_key)
                except sr.UnknownValueError:
                    ind_text = ""
                    #shutil.copyfile(os.path.join(self.mp3_path, v), os.path.join(self.mp3_path, "NotRecognized", v))
                except sr.RequestError as e:
                    ind_text = ""
                    #shutil.copyfile(os.path.join(self.mp3_path, v), os.path.join(self.mp3_path, "NotRecognized", v))
                converted_list.append(ind_text)
            text = "".join(converted_list)
        return text


    def submit(self):
        if self.entry1.get() == "":
            messagebox.showerror("Error", "Please provide the path to media files")
            self.button1.config(state="enabled")
        elif self.entry2.get() == "":
            messagebox.showerror("Error", "Please provide the path to Raw Data file")
            self.button1.config(state="enabled")
        elif not('\\' not in self.entry1.get() or "/" not in self.entry1.get()):
            messagebox.showerror("Error", "Please provide correct path to Media")
            self.button1.config(state="enabled")
        elif ".xlsx" not in self.entry2.get():
            messagebox.showerror("Error", "Please provide correct path to Raw Data file")
            self.button1.config(state="enabled")
        else:
            self.mp3_path = os.path.normpath(self.entry1.get())
            excel_path = os.path.normpath(self.entry2.get())

            #os.mkdir(os.path.join(self.mp3_path, "NotRecognized"))

            files = os.listdir(self.mp3_path)
            mpeg4_files = [x for x in files if '.wav' in x]


            wb1 = openpyxl.load_workbook(excel_path, keep_vba=True)
            wb2 = openpyxl.Workbook()
            sheet1 = wb1["Sheet1"]
            sheet2 = wb2.active
            sheet2.cell(row=1, column=1).value = "Respondent Serial No"
            sheet2.cell(row=1, column=2).value = "Centre"
            sheet2.cell(row=1, column=3).value = "Respondent Name"
            sheet2.cell(row=1, column=4).value = "Language"

            row_num = 2
            j = True
            file_num = 1

            while (j):
                intnr = int(sheet1.cell(row=row_num, column=1).value)
                sheet2.cell(row=row_num, column=1).value = intnr
                sheet2.cell(row=row_num, column=2).value = sheet1.cell(row=row_num, column=2).value
                sheet2.cell(row=row_num, column=3).value = sheet1.cell(row=row_num, column=5).value
                sheet2.cell(row=row_num, column=4).value = sheet1.cell(row=row_num, column=7).value

                file_name = [x for x in mpeg4_files if intnr == int(x[:8])]
                u = 0
                col_num = 0

                for u,v in enumerate(file_name):
                    updated_file = [x for x in file_name if x[17:22] == v[17:22]]
                    if v != updated_file[0]:
                        file_num += 1
                        continue
                    else:
                        self.pbar['value'] = 100 * (file_num) / len(mpeg4_files)
                        self.name.set(f"Processing File: {v[:-8]}")
                        self.root.update()

                        text = self.combined_text(updated_file)

                        if text == "":
                            text = "Audio Not Recognized"

                        if row_num == 2:
                            sheet2.cell(row=1, column=col_num * 2 + 5).value = "File Name"
                            sheet2.cell(row=1, column=col_num * 2 + 6).value = "Text"
                            sheet2.cell(row=row_num, column=col_num * 2 + 5).value = v.replace(".wav", "")
                            path = f'{os.path.join(self.mp3_path.replace("FinalMedia", ""), v[:-8])}.wav'
                            sheet2.cell(row=row_num, column=col_num * 2 + 5).hyperlink = path
                            sheet2.cell(row=row_num, column=col_num * 2 + 6).value = text
                        else:
                            sheet2.cell(row=row_num, column=col_num * 2 + 5).value = v.replace(".wav", "")
                            path = f'{os.path.join(self.mp3_path.replace("FinalMedia", ""), v[:-8])}.wav'
                            sheet2.cell(row=row_num, column=col_num * 2 + 5).hyperlink = path
                            sheet2.cell(row=row_num, column=col_num * 2 + 6).value = text

                        #os.remove(os.path.join(self.mp3_path, v))
                        file_num += 1
                        col_num += 1
                    

                row_num += 1
                if sheet1.cell(row=row_num, column=1).value == None:
                    j = 0

            shutil.rmtree(os.path.join(self.mp3_path))
            wb2.save("Output.xlsx")
            wb1.close()
            messagebox.showinfo("Information", "Completed")
            self.button1.config(state="enabled")

    def thread_func(self):
        self.button1.config(state="disabled")
        self.thread = threading.Thread(target=self.submit)
        self.thread.setDaemon(True)
        self.thread.start()


if __name__ == "__main__":
    root = ThemedTk(theme="radiance")
    myapp = MyApp(root)
    root.mainloop()
        
