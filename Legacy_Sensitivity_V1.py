# %%
import ttkbootstrap as ttkbst
from ttkbootstrap.constants import *

import tkinter.scrolledtext as st

import threading
import Function as func

Win_GUI = ttkbst.Window(title="LTE Sensitivity Drawing", themename="cosmo")
Win_GUI.attributes("-topmost", True)
Win_GUI.geometry("865x510")

s = ttkbst.Style()
s.configure("MIPI.TButton", font=("Calibri", 20, "bold"))

btn_add_file = ttkbst.Button(
    Win_GUI,
    text="Draw (F5)",
    style="MIPI.TButton",
    command=lambda: [
        func.add_file(Entry_file_path),
        threading.Thread(target=func.LTE_Sens_drawing, args=(Win_GUI, Entry_file_path, text_area)).start(),
    ],
)
btn_add_file.place(x=5, y=40, width=300, height=100)

Entry_file_path = ttkbst.Entry(Win_GUI, width=60)
Entry_file_path.insert(
    0,
    "C:\Labtest\Report\Spring",
)
Entry_file_path.place(x=5, y=5, width=300, height=30)

Scrolled_txt_frame = ttkbst.Frame(Win_GUI)
Scrolled_txt_frame.place(x=310, y=5, width=550, height=500)

text_area = st.ScrolledText(Scrolled_txt_frame, font=("Consolas", 9))
text_area.place(x=0, y=0, width=550, height=500)

Win_GUI.bind(
    "<F5>",
    lambda event: [
        func.add_file(Entry_file_path),
        threading.Thread(target=func.LTE_Sens_drawing, args=(Win_GUI, Entry_file_path, text_area)).start(),
    ],
)


def Win_GUI_close():
    Win_GUI.quit()


Win_GUI.resizable(False, False)
Win_GUI.focus()
Win_GUI.mainloop()
