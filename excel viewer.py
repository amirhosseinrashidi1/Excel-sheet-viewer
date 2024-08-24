from tkinter import *
from tkinter import messagebox
from tkinter import ttk, filedialog
import pandas as pd

root = Tk()
root.title("Excel DataSheet Viewer")
root.geometry('1200x600+200+100')

def open_file():
    filename = filedialog.askopenfilename(title="Open a file", 
                                          filetypes=(("Excel Files", "*.xlsx"),
                                                     ("All files", "*.*")))
    if filename:
        try:
            df = pd.read_excel(filename)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open the file.\n{str(e)}")
            return

        tree.delete(*tree.get_children())

        tree['columns'] = list(df.columns)
        tree['show'] = "headings"

        for col in tree['columns']:
            tree.heading(col, text=col)
            tree.column(col, width=150)

        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            tree.insert("", "end", values=row)

        tree.yview_moveto(0)
        tree.xview_moveto(0)

try:
    image_icon = PhotoImage(file="logo.png")
    root.iconphoto(False, image_icon)
except Exception as e:
    print("Icon image could not be loaded.")

frame = Frame(root)
frame.pack(pady=20, fill=BOTH, expand=True)

tree_scroll_x = Scrollbar(frame, orient=HORIZONTAL)
tree_scroll_x.pack(side=BOTTOM, fill=X)
tree_scroll_y = Scrollbar(frame, orient=VERTICAL)
tree_scroll_y.pack(side=RIGHT, fill=Y)

tree = ttk.Treeview(frame, xscrollcommand=tree_scroll_x.set, yscrollcommand=tree_scroll_y.set)
tree.pack(fill=BOTH, expand=True)

tree_scroll_x.config(command=tree.xview)
tree_scroll_y.config(command=tree.yview)

button = Button(root, text='Open', width=60, height=2, font=30, fg="white", bg="#0078d7", 
                command=open_file)
button.pack(padx=10, pady=10)

root.mainloop()
