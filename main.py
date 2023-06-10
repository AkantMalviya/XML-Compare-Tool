from tkinter import *
import os
from BrowseFile import *
from CompareXML import *

#Variable declaration
global file_path

# Fonts
copyryt = u"\u00A9"
trademark = u"\u2122"
font4 = ('Times', 18)
font1 = ('Times', 15, 'bold')
font2 = ('Times', 12)
font3 = ('Times', 11, 'bold')
brand = copyryt + 'AkantMalviya'

# Windows initialization
root = Tk()
root.title("XML Compare Tool")
root.option_add("*tearOff", False)
root.resizable(0,0)
photo = PhotoImage(file='xml.png')
root.iconphoto(False, photo)
window_height = 220
window_width = 580
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width / 2) - (window_width / 2))
y_cordinate = int((screen_height / 2) - (window_height / 2))
root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

def show_about_info():
    messagebox.showinfo(
        title="About",
        message=f'''XML Compare Tool\n\nPlease Select Backup XML File and Updated XML File from Browse and Press Compare>>\n\nFor any modification & help, please contact {brand}''') #


def openBackupFileBrowser(txt_backup):
    file_path = browseFile()
    txt_backup.delete(1.0, "end")
    txt_backup.insert(1.0, file_path)


def openUpdatedFileBrowser(txt_updated):
    file_path = browseFile()
    txt_updated.delete(1.0, "end")
    txt_updated.insert(1.0, file_path)

def Refresh():
    txt_backup.delete(1.0, "end")
    txt_updated.delete(1.0, "end")


def location():
    folderpath = os.path.join(os.getcwd(), 'CompareResults')
    os.startfile(folderpath)


# Defining the Widgets
backup_label = Label(root, text="Backup File    ", font=font3)
updated_label = Label(root, text="Updated File   ", font=font3)
txt_backup = Text(root, bd=3, heigh=0, width=35)
txt_backup.insert(1.0,"")
txt_updated = Text(root, bd=3, height=0, width=35)
txt_updated.insert(1.0,"")
button_browse1 = Button(root, text="Browse", font= font3, padx=20, pady=5, command=lambda: openBackupFileBrowser(txt_backup))
button_browse2 = Button(root, text="Browse", font= font3, padx=20, pady=5, command=lambda: openUpdatedFileBrowser(txt_updated))
button_compare = Button(root, text="COMPARE", font= font3, padx=20, pady=5, command=lambda: compare_xml_files(txt_backup, txt_updated))
# button_ok = Button(root, text="Exit", padx=20, pady=5, command=root.quit)
# button_cancel = Button(root, text="Cancel", padx=20, pady=5, command=root.quit)


# Menubar Options Help
font5 = ('Times', 12, 'bold')

menubar = Menu()

root.config(menu=menubar)
options_menu = Menu(menubar)
help_menu = Menu(menubar)

menubar.add_cascade(menu=options_menu, label="Options")
menubar.add_cascade(menu=help_menu, label="Help")

options_menu.add_command(label="Refresh", command=lambda: Refresh())
options_menu.add_command(label="Location", command=lambda: location())
options_menu.add_command(label="Exit", command= lambda: root.quit())
help_menu.add_command(label="Instruction", command=lambda: show_about_info())

# image logo
img = PhotoImage(master=root, file="logo.png")
img = img.subsample(7,7)
logo = Label(master=root, image=img)

# Positioning the Widgets
backup_label.grid(row=0, column=1, padx= 10 ,pady=10 ,ipadx=10 ,ipady=10)
txt_backup.grid(row=0, column=2, columnspan=3, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
button_browse1.grid(row=0, column=6, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
updated_label.grid(row=1, column=1, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
txt_updated.grid(row=1, column=2, columnspan=3, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
button_browse2.grid(row=1, column=6, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)

button_compare.grid(row=2, column=3, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
logo.grid(row=2, column=1, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
#button_ok.grid(row=2, column=4, padx= 10 ,pady=10 ,ipadx=1 ,ipady=1)
#button_cancel.grid(row=5, column=6)

root.mainloop()





