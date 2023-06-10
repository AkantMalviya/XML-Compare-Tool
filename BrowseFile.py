from tkinter import *
from tkinter import filedialog

def browseFile():
    root = Tk()
    root.withdraw()
    root.filename = filedialog.askopenfilename(title="Select a File",
                                               filetype=(("XML Files", "*.xml"), ("All Files", "*.*")))
    return root.filename

