
def excelsheet():
    s=e.get()
    import xlwt 
    from xlwt import Workbook 
    ss=z.get()
    zz=str(ss).replace("\\","\\\\")
    wb = Workbook() 
    sheet1 = wb.add_sheet('Sheet 1') 
    sheet1.write(1, 0, 'ISBT DEHRADUN') 
    sheet1.write(2, 0, 'SHASTRADHARA') 
    sheet1.write(3, 0, 'CLEMEN TOWN') 
    sheet1.write(4, 0, 'RAJPUR ROAD') 
    wb.save(zz+'\\'+s+'.xls')
    e.delete(0,END)

    

  
    

from tkinter import *

root = Tk()
root.title("Counting Seconds")
root.geometry("210x100")
root.resizable(False,False)
Label(root,text = 'Enter File Path').grid(row=0)
z = Entry(root)
z.grid(row=0,column=1)
z.focus_set()


Label(root,text = 'Enter File Name').grid(row=1)
e = Entry(root)
e.grid(row=1,column=1)
e.focus_set()

button = Button(root, text='Create Excel', command = excelsheet)
button.grid(row=3)

Button(root,text='Copy to Excel').grid(row=3,column=1)

Button(root,text='Copy to Excel').grid(row=3,column=2)


root.mainloop()
