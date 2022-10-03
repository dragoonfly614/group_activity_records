'''
Group buiding page:
Read an excel file to extract student names and construct a button for each student without a group.

Assign group members: press the group number button then student name button will record the student as a member and then remove the student name button from the grid.
    Display the student name as a button under the gorup name button.

Delete group member: click student name button in a group to unbind.

Assign group leader: Double-click stdent name button to disable old leader, if applicable, and assign the student as leader with a special background color.

Group adding bonus points: Double-click the group name button to display (in another page) the current bonus point status,two entry to enter week and points for the group and a submit button to submit.


'''

import tkinter as tk
import openpyxl



def click(event):
    print(event.widget.cget('text'))
    #buttons[0].config(text = event.char)

    #event.widget.destroy()
    #btn_1.pack()

def onclick(event,i):
    print(i)
    event.widget.grid_forget()
    print(event.widget.cget('text'))
    student_name = event.widget.cget('text')
    global group_index
    group_buts[group_index].append(tk.Button(frame_1,text = student_name, width=10))
    group_buts[group_index][-1].pack()

def group1_choice(event):
    global group_index
    group_index = 0

def group_bonus_points(event):
    print('DC')
    pass

def group_leader(event):
    print('RC')
    pass
# this part should be separate as an individual module
my_excel_file = openpyxl.load_workbook('attendence.xlsx')
sheets = my_excel_file.sheetnames

sheet = my_excel_file[sheets[0]]

#print(sheets,sheet)
#week = int(input('周次：'))
#column = chr(ord('E')+week-1)

name_row = 5


window = tk.Tk()
frame_1 = tk.Frame(window, relief = tk.RAISED, bd = 2)
frame_2 = tk.Frame(window, relief = tk.RAISED, bd = 2)
frame_1.grid(row = 1, column = 1)
frame_2.grid(row = 1, column = 2)

btn_1 = tk.Button(frame_1,text = 'abc', width=10)
btn_1.bind('<Button-1>',group1_choice)
btn_1.bind('<Double-Button-1>',group_bonus_points)
#btn_1.bind('<Button-3>',group_leader)
btn_1.pack()

buttons = []
group_buts = [[] for i in range(7)]

group_index = 0

i = 0

while sheet['C%d'%name_row].value != None:
    student_name = sheet['C%d'%name_row].value
    buttons.append(tk.Button(frame_2,text = student_name, width=10))
    buttons[i].bind('<Button-1>',lambda event, i=1: onclick(event, i))
    buttons[i].grid(row = i//3+1, column =i%3+1)
    #print(f'i={i}, name_row={name_row}, name={student_name}')
    i += 1
    name_row += 1


window.mainloop()
