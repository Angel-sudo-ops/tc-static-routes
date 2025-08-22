import tkinter as tk
import os

root = tk.Tk()
root.title("Test Window")

icon_path = os.path.abspath("route.ico")
print(f"Icon path: {icon_path}")
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)
else:
    print("Icon file not found.")


root.geometry("470x480")
# root.minsize(470, 480)
# root.resizable(True, True)
root.mainloop()