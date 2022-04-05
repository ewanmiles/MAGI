import tkinter as tk
import sys

class LogWindow(tk.Tk):

    def __init__(self, windowGeo: str, frameGeo: tuple):
        """
        Classic OOP init for tkinter window, args:
            - windowGeo (str): Geometry for the window, e.g. '500x500'
            - frameGeo (tuple<int, int>): Geometry for the log frame (textarea), given as (width, height), e.g. (400, 400)
        """
        super().__init__() #Calls the init for the Tk class

        self.geometry(windowGeo)
        self.frameWidth = frameGeo[0] #Both accessed by log functions
        self.frameHeight = frameGeo[1]

        self.acronym = tk.Label(text="MAGI", pady=5, font=('Felix Titling', 45))
        self.title = tk.Label(text="Madcap Automatic Graphic Implementation v2.0", pady=5, font=('Footlight MT Light', 15))

        self.log = tk.Frame(
            master=self, 
            width=self.frameWidth, 
            height=self.frameHeight, 
            bg="white"
        )
        
        self.info = tk.Label(
            master=self.log, 
            text="Starting the process...", 
            bg="white", 
            justify='left', 
            wraplength=self.frameWidth-30
        )
        self.info.place(x=0, y=0) #Text starts in top left of frame

        self.button = tk.Button(
            text="Cancel",
            width=25,
            height=2,
            bg="azure",
            command=sys.exit,
        )

        for el in [self.acronym, self.title, self.log]:
            el.pack()

        self.button.pack(pady=40)

    def setButtonText(self, text):
        #Not necessary accessor but easier to read
        self.button['text'] = text

    def logText(self, text: str):
        """
        Really basic. Just a print statement to the console that will also update the tkinter window text.
        text is Stringtype only.
        """

        if self.info.winfo_height() > self.frameHeight - 10: #If the frame height has been exceeded by text, clear all previous text
            self.info['text'] = ''

        #Update console and window
        print(f'\n{text}')
        self.info['text'] += f'\n{text}'