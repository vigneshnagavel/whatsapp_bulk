import tkinter as tk
import tkinter.font as tkFont
import pywhatkit as w
import time
import pyautogui
import keyboard as k
import pandas as pd
import art
from datetime import datetime

data = []

class App:
    def __init__(self, root):
        #setting title
        root.title("Whatsapp Auto")
        #setting window size
        width=800
        height=600
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        messagetext = tk.StringVar()  
        path = tk.StringVar() 

        GLabel_541=tk.Label(root)
        GLabel_541["anchor"] = "center"
        GLabel_541["bg"] = "#c6c6c6"
        ft = tkFont.Font(family='Times',size=10)
        GLabel_541["font"] = ft
        GLabel_541["fg"] = "#333333"
        GLabel_541["justify"] = "center"
        GLabel_541["text"] = "Enter Message"
        GLabel_541.place(x=110,y=150,width=110,height=50)

        GLabel_421=tk.Label(root)
        GLabel_421["bg"] = "#195bcd"
        ft = tkFont.Font(family='Times',size=13)
        GLabel_421["font"] = ft
        GLabel_421["fg"] = "#ffffff"
        GLabel_421["justify"] = "center"
        GLabel_421["text"] = "Whatsapp Auto Message Sender"
        GLabel_421.place(x=260,y=30,width=282,height=59)

        GLineEdit_689=tk.Entry(root)
        GLineEdit_689["borderwidth"] = "2px"
        ft = tkFont.Font(family='Times',size=10)
        GLineEdit_689["font"] = ft
        GLineEdit_689["fg"] = "#000000"
        GLineEdit_689["justify"] = "center"
        GLineEdit_689["text"] = "Available soon!"
        GLineEdit_689.place(x=240,y=150,width=438,height=194)

        GLabel_651=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_651["font"] = ft
        GLabel_651["fg"] = "#333333"
        GLabel_651["justify"] = "center"
        GLabel_651["text"] = "By ECE Department"
        GLabel_651.place(x=320,y=90,width=157,height=30)

        GLabel_253=tk.Label(root)
        GLabel_253["bg"] = "#cdcdcd"
        ft = tkFont.Font(family='Times',size=10)
        GLabel_253["font"] = ft
        GLabel_253["fg"] = "#333333"
        GLabel_253["justify"] = "center"
        GLabel_253["text"] = "Enter Excel Path"
        GLabel_253.place(x=110,y=360,width=110,height=50)

        GLineEdit_633=tk.Entry(root)
        GLineEdit_633["borderwidth"] = "3px"
        ft = tkFont.Font(family='Times',size=10)
        GLineEdit_633["font"] = ft
        GLineEdit_633["fg"] = "#000000"
        GLineEdit_633["justify"] = "center"
        GLineEdit_633["text"] = "Available soon!"
        GLineEdit_633.place(x=240,y=360,width=436,height=50)

        

        GButton_97=tk.Button(root)
        GButton_97["bg"] = "#0b961e"
        ft = tkFont.Font(family='Times',size=18)
        GButton_97["font"] = ft
        GButton_97["fg"] = "#ffffff"
        GButton_97["justify"] = "center"
        GButton_97["text"] = "Send"
        GButton_97.place(x=300,y=470,width=212,height=64)
        GButton_97["command"] = self.GButton_97_command

        GLabel_691=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_691["font"] = ft
        GLabel_691["fg"] = "#333333"
        GLabel_691["justify"] = "center"
        GLabel_691["text"] = "Note - Make sure that each time the max contacts be 300 thus your whatsapp account not get banned"
        GLabel_691.place(x=110,y=560,width=588,height=30)


        global FILE_LOC 
        global NAME 
        global PHONENUMBER 
        global SENT_DATA
        FILE_LOC = "D:\programming\whatsappauto\data.xlsx"
        NAME = 'Name'
        PHONENUMBER = 'Phone Number'
        SENT_DATA = "D:\programming\whatsappauto\sentdata.txt"
        
        
    def GButton_97_command(self):
        print(FILE_LOC)
        def read_data_from_excel():
            try:
                df = pd.read_excel(FILE_LOC)
                print("Retrieving data from excel", 'INFO')
            except:
                print("Excel 'data.xlsx' not found", 'ERROR')
            print("Found {0} messages to be send".format(len(df.index)), 'INFO')
            for i in df.index:
                if '+' not in str(df[PHONENUMBER][i]):
                    number = '+91' + str(df[PHONENUMBER][i])
                else:
                    number = str(df[PHONENUMBER][i])
                output = {
                    
                    'Name': df[NAME][i],
                    'PhoneNumber': number,
                    'Message': "hi"
                }
                data.append(output)

        read_data_from_excel()
        count = 0
        for i in range(len(data)):
            w.sendwhatmsg_instantly(data[i]["PhoneNumber"], """*VSB ENGINEERING COLLEGE*
*KARUR & COVAI* 🎓🎓🎓🎓
+2 மாணவர்களே எங்கள் கல்லூரியில் Engineering course சேர்க்கை நடைபெறுகிறது. விருப்பம் உள்ள மாணவர்கள் தொடர்பு கொள்ளவும்.

*COURSES OFFERED*
1.Mechanical Engg
2.Civil Engg
3.ECE
4.EEE
5.AIDS
6.CSE
7.Bio Medical
8.Bio Technology
9.Chemical
10. Agriculture Engg
11.Information technology
12.Computer Communication and Engg
13.Computer Science and Business System


*குறிப்பு*: 180 மற்றும் அதற்கு மேல் மதிபெண் பெற்றவர்களுக்கும், 190 மற்றும் அதற்கு மேல் மதிப்பெண் பெற்றவர்களுக்கும் கல்லூரி நிர்வாகம் கட்டண சலுகை அறிவித்துள்ளது. மேலும் விவரங்களுக்கு

*S.ராஜேஷ் குமார்*,
Assistant Professor
Department of ECE
7010379002
8825540246
💐 welcome to all💐*Feel Free to check out Brochure!*
https://vsbecinstitution.github.io/RAJESHKUMAR_ECE.pdf""", 15)
            count = count + 1
            print(art.text2art(str(count)))
            pyautogui.click(800, 800)
            time.sleep(2)
            pyautogui.click(800, 800)
            k.press_and_release('enter')

            time.sleep(3)
            k.press_and_release('ctrl+w')

            now = datetime.now()

            current_time = now.strftime("%H:%M:%S")
            more_lines = [data[i]["Name"],"\t",data[i]["PhoneNumber"],"\t",current_time,"\n"]
            with open(SENT_DATA, 'a') as f:
                f.writelines(more_lines) # closes the last tab
    

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
