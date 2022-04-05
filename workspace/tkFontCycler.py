import tkinter as tk
import tkinter.font as tkFont

root = tk.Tk()
fonts = list(tkFont.families())
j = 0
for i in fonts[j:j+10]:
    tk.Label(root, text=f"MAGI - {i}", pady=5, font=(i, 25)).pack()

def nextFonts():
    global j 
    j += 10

    for widgets in root.winfo_children():
        widgets.destroy()

    for i in fonts[j:j+10]:
        tk.Label(root, text=f"MAGI - {i}", pady=5, font=(i, 25)).pack()

    button = tk.Button(
            text="Next",
            width=25,
            height=2,
            bg="azure",
            command=nextFonts,
        )
    button.pack()

button = tk.Button(
            text="Next",
            width=25,
            height=2,
            bg="azure",
            command=nextFonts,
        )
button.pack()

root.mainloop()