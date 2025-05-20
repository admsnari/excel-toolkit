import ttkbootstrap as ttk
from ttkbootstrap.constants import *

root = ttk.Window()
root.title("Main Window")
root.geometry("300x120")

def open_progress_window():
    # Create new window
    progress_win = ttk.Toplevel(root)
    progress_win.title("Progress Window")
    progress_win.geometry("300x100")

    # Create and pack progress bar
    progress = ttk.Progressbar(progress_win, length=250, mode='determinate', bootstyle=SUCCESS)
    progress.pack(pady=20)

    # Function to simulate progress
    def start_progress(value=0):
        if value <= 100:
            progress['value'] = value
            progress_win.after(50, start_progress, value + 5)
        else:
            progress.stop()

    start_progress()

# Main button to trigger progress window
b1 = ttk.Button(root, text="Show Progress", bootstyle=SUCCESS, command=open_progress_window)
b1.pack(pady=20)

root.mainloop()
