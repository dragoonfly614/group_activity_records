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

# constants
CLASS_1 = 'points_ele1.xlsx'
CLASS_2 = 'points_rob1.xlsx'
CLASS_3 = 'points_rob2.xlsx'

# find the name in the sheet then return row number

# write the group number into the corresponding cell w.r.t. name and sheet
def add_group(name, sheet, group_index, my_excel_file, excel_path):
    global name2row_dict
    sheet['E%d'%name2row_dict[name]].value = group_index
    my_excel_file.save(excel_path)

def remove_group(name, sheet, my_excel_file, excel_path):
    global name2row_dict
    sheet['E%d'%name2row_dict[name]].value = None
    sheet['F%d'%name2row_dict[name]].value = None
    my_excel_file.save(excel_path)

def leader_choice(event):
    global group_edit_flag, name2row_dict
    name = event.widget.cget('text')
    leader_flag = sheet['F%d'%name2row_dict[name]].value
    if group_edit_flag:
        if leader_flag == None:
            sheet['F%d'%name2row_dict[name]].value = 1
            my_excel_file.save(excel_path)
            event.widget.config(bg = 'green',fg = 'white')
        else:
            sheet['F%d'%name2row_dict[name]].value = None
            my_excel_file.save(excel_path)
            event.widget.config(bg = 'SystemButtonFace',fg = 'black')

# click action of frame_1 group members
def group_click(event):
    global group_edit_flag
    if group_edit_flag:
        event.widget.grid_forget()
        buttons.append(tk.Button(frame_2,text = event.widget.cget('text'), width = 10))
        buttons[-1].grid(row = (len(buttons)- 1) // 3 + 1, column = (len(buttons) - 1) % 3 + 1)
        remove_group(event.widget.cget('text'), sheet, my_excel_file, excel_path)

# click action of frame_2 students(without a group)
def onclick(event):
    global group_edit_flag
    if group_edit_flag:
        event.widget.grid_forget()
        print(event.widget.cget('text'))
        student_name = event.widget.cget('text')
        global group_index_choice
        group_buts[group_index_choice].append(tk.Button(frame_1,text = student_name, width = 10))
        group_buts[group_index_choice][-1].bind('<Button-1>', group_click)
        group_buts[group_index_choice][-1].bind('<Button-3>', leader_choice)
        # place the button in the block(3 rows) of group_name, the 2nd row or 3rd row(1st rowfor group_name), and corresponding column(4 column a row)
        group_buts[group_index_choice][-1].grid(row = group_index_choice * 3 + (len(group_buts[group_index_choice]) - 1) // 4 + 2, column = (len(group_buts[group_index_choice]) - 1) % 4 + 1)
        add_group(student_name, sheet, group_index_choice, my_excel_file, excel_path)

# click action of group name buttons
def group_choice(event, i):
    global group_edit_flag
    group_edit_flag = not(group_edit_flag)
    if group_edit_flag:
        event.widget.config(bg = 'red',fg = 'white')
    else:
        event.widget.config(bg = 'SystemButtonFace',fg = 'black')
    global group_index_choice
    #print(f'group_index_choice={group_index_choice}, i={i}')
    group_index_choice = i


def group_bonus_points(event):
    pass


# this part should be separate as an individual module
class_choice = input('1 for Applied Electronic major class one, 2 for Robot major class one， 3 for Robot major class two: ')
if class_choice == '1':
    excel_path = CLASS_1
elif class_choice == '2':
    excel_path = CLASS_2
elif class_choice == '3':
    excel_path = CLASS_3

my_excel_file = openpyxl.load_workbook(excel_path)
sheets = my_excel_file.sheetnames

sheet = my_excel_file[sheets[0]]
name2row_dict = {}

#print(sheets,sheet)
#week = int(input('周次：'))
#column = chr(ord('E')+week-1)

name_row = 5
group_edit_flag = False
group_index_choice = 0


window = tk.Tk()
frame_1 = tk.Frame(window, relief = tk.RAISED, bd = 2)
frame_2 = tk.Frame(window, relief = tk.RAISED, bd = 2)
frame_1.grid(row = 1, column = 1)
frame_2.grid(row = 1, column = 2)

buttons = []
group_name_buts = []
group_buts = [[] for i in range(7)]

for index in range(7):
    group_name_buts.append(tk.Button(frame_1,text = f'Group{index}', width=10, bd = 3))
    group_name_buts[index].bind('<Button-1>',lambda event, i = index: group_choice(event, i))
    group_name_buts[index].bind('<Double-Button-1>',group_bonus_points)
    #btn_1.bind('<Button-3>',group_leader)
    group_name_buts[index].grid(row = index*3+1,column = 1)


while sheet['C%d'%name_row].value != None:
    student_name = sheet['C%d'%name_row].value
    name2row_dict[student_name] = name_row
    if sheet['E%d'%name_row].value == None:
        buttons.append(tk.Button(frame_2,text = student_name, width = 10))
        buttons[-1].bind('<Button-1>', onclick)
        buttons[-1].grid(row = (len(buttons) - 1) // 3 + 1, column = (len(buttons) - 1) % 3 + 1)
    #print(f'i={i}, name_row={name_row}, name={student_name}')
    else:
        group_number = sheet['E%d'%name_row].value
        group_buts[group_number].append(tk.Button(frame_1,text = student_name, width = 10))
        group_buts[group_number][-1].bind('<Button-1>', group_click)
        group_buts[group_number][-1].bind('<Button-3>', leader_choice)
        # place the button in the block(3 rows) of group_name, the 2nd row or 3rd row(1st rowfor group_name), and corresponding column(4 column a row)
        group_buts[group_number][-1].grid(row = group_number * 3 + (len(group_buts[group_number]) - 1) // 4 + 2, column = (len(group_buts[group_number]) - 1) % 4 + 1)
        if sheet['F%d'%name_row].value == 1:
            group_buts[group_number][-1].config(bg = 'green',fg = 'white')
    name_row += 1


window.mainloop()
