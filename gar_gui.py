'''
Group buiding page:
Read an excel file to extract student names and construct a button for each student without a group.

Edit group members:

'''

import tkinter as tk

def func(event):
    btn_1.config(text = event.widget.get())
    event.widget.destroy()
    btn_1.pack()

def onclick():
    btn_1.pack_forget()
    ent_1 = tk.Entry(frame)
    ent_1.pack()
    ent_1.bind('<Return>',func)
    ent_1.focus_set()


window = tk.Tk()
frame = tk.Frame(window, relief = tk.RAISED, bd = 2)
btn_1 = tk.Button(frame,text = '',command = onclick)
btn_1.pack()
frame.pack()
window.mainloop()
