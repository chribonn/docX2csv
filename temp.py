from tkinter import *
import ttkbootstrap as tb

root =tb.Window(themename="superhero")

def clicker():
    my_label.config(text=f"You clicked on {my_combo.get()}")
    
def click_bind(e):
    my_label.config(text=f"You clicked on {my_combo.get()} !")
    
 
root.title("TTK Bootstrap! Combobox")
root.geometry('500x350')

my_label = tb.Label(root, text="Hello World!", font=("Helvetica", 18))
my_label.pack(pady=30)

# Create dropdown options
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Satuday", "Sunday"]

my_combo = tb.Combobox(root, bootstyle="success", values=days)
my_combo.pack(pady=30)

#set combo  default
my_combo.current(0)

# Create button
mybutton = tb.Button(root, text="Click Me!", command=clicker, bootstyle="danger")
mybutton.pack(pady=20)

# Bind the combobox
my_combo.bind("<<ComboboxSelected>>", click_bind)

root.mainloop()