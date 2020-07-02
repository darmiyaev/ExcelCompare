filenameone = ""
filenametwo = ""
from itertools import zip_longest
import xlrd
from tkinter import *
from tkinter import filedialog
def start():
    global filenameone
    global filenametwo
start()
def browseFilesone():
    print(1)
    global filenameone
    filenameone = filedialog.askopenfilename(initialdir = "/", 
                                          title = "Select a File", 
                                          filetypes = (("Excel (.xlsx)", 
                                                        "*.xlsx*"), 
                                                       ("All files", 
                                                        "*.*")))
    print(filenameone)
def browseFilestwo():
    print(2)
    global filenametwo
    filenametwo = filedialog.askopenfilename(initialdir = "/", 
                                          title = "Select a File", 
                                          filetypes = (("Excel (.xlsx)", 
                                                        "*.xlsx*"), 
                                                       ("All files", 
                                                        "*.*")))
    print(filenametwo)
def check():
    print("check")
    print(filenameone)
    print(filenametwo)
    rb1 = xlrd.open_workbook(filenameone)
    rb2 = xlrd.open_workbook(filenametwo)
    sheet1 = rb1.sheet_by_index(0)
    sheet2 = rb2.sheet_by_index(0)
    pos = 0
    new_label = Text(window, wrap="word")
    scrollbar = Scrollbar(window, command=new_label.yview)
    new_label.configure(yscrollcommand=scrollbar.set)
    for rownum in range(max(sheet1.nrows, sheet2.nrows)):
        if rownum < sheet1.nrows:
            row_rb1 = sheet1.row_values(rownum)
            row_rb2 = sheet2.row_values(rownum)
            for colnum, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)):
                if c1 != c2:
                    print("Row {} Col {} - {} != {}".format(rownum+1, colnum+1, c1, c2))
                    pos = pos + 40
                    message = "Row {} Col {} - {}  is not equal to {}".format(rownum+1, colnum+1, c1, c2)
                    new_label.insert("end", message+"\n")
                    new_label.grid(column = 1, row = pos)
                    #new_label = Label(text="Row {} Col {} - {} is not equal to {}".format(rownum+1, colnum+1, c1, c2))
                #else:
                        #print("Row {} missing".format(rownum+1))

    # Change label contents 
    #label_file_explorer.configure(text="File Opened: "+ filenameone + filenametwo) 
       

                                                                                                   
# Create the root window 
window = Tk() 
#my_label = Label(text=)
# Set window title 
window.title('ExcelCompare') 
   
# Set window size 
window.geometry("500x500") 
   
#Set window background color 
window.config(background = "white") 
   
# Create a File Explorer label 
label_file_explorer = Label(window,  
                            text = "ExcelCompare", 
                            width = 100, height = 4,  
                            fg = "blue") 
   
button_first = Button(window,  
                        text = "Excel file 1", 
                        command = browseFilesone)
button_second = Button(window,  
                        text = "Excel file 2", 
                        command = browseFilestwo)  
'''
button_exit = Button(window,  
                     text = "Exit", 
                     command = exit)  
   '''
# Grid method is chosen for placing 
# the widgets at respective positions  
# in a table like structure by 
# specifying rows and columns 
label_file_explorer.grid(column = 1, row = 1) 

#my_label.grid(column = 1, row = 11)
button_first.grid(column = 1, row = 2)
button_second.grid(column = 1, row = 7) 
   
#button_exit.grid(column = 1,row = 3)
button_check = Button(window,  
                        text = "Check", 
                        command = check)
button_check.grid(column = 1, row = 9) 
# Let the window wait for any events 
window.mainloop() 
