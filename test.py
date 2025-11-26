# Source - https://stackoverflow.com/a
# Posted by Ed Ward, modified by community. See post 'Timeline' for change history
# Retrieved 2025-11-26, License - CC BY-SA 4.0

import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
root = tk.Tk()
def change_i():
    if sound_btn.image == icon:
        #start_recording()

        sound_btn.config(image=icon)
        sound_btn.image = icon
    else:
        #stop_recording()

        sound_btn.config(image=icon)
        sound_btn.image = icon

icon = PhotoImage(file='icons/print.png')
# icon2 = PhotoImage(file='stop.png')

sound_btn = tk.Button(root, image=icon, width=70,height=60,relief=FLAT ,command=change_i )
sound_btn.image = icon
sound_btn.grid(row=0, column=1)
root.mainloop()
