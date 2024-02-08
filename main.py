import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx2pdf import convert


def reset():
    canvas1.delete("all")  # Clear canvas
    initialize_widgets()


def initialize_widgets():
    global canvas1, button1, label1, label2, progress_bar

    canvas1 = tk.Canvas(root, width=400, height=350, relief='raised')
    canvas1.pack()

    root.iconbitmap(default="icon.png")

    label1 = tk.Label(root, text='Doc to Pdf Converter')
    label1.config(font=('helvetica', 14))
    canvas1.create_window(200, 45, window=label1)

    label2 = tk.Label(root, text='Choose Your Docx Files :')
    label2.config(font=('helvetica', 10))
    canvas1.create_window(200, 100, window=label2)

    progress_bar = ttk.Progressbar(root, maximum=100)
    canvas1.create_window(200, 180, window=progress_bar)

    button1 = tk.Button(text='Browse Your Files', command=path_name, bg='grey', fg='black',
                        font=('helvetica', 9, 'bold'))
    canvas1.create_window(200, 150, window=button1)


def path_name():
    paths_files = filedialog.askopenfilenames()

    button2 = tk.Button(text='Convert', command=lambda p=paths_files: converted(p), bg='brown', fg='white',
                        font=('helvetica', 9, 'bold'))
    canvas1.create_window(200, 230, window=button2)


def converted(paths_files):
    total_files = len(paths_files)
    increment_value = 100 / total_files
    current_progress = 0

    for i, path in enumerate(paths_files, start=1):
        convert(path)
        current_progress += increment_value
        progress_bar["value"] = current_progress
        root.update_idletasks()

    messagebox.showinfo("Conversion Complete", "Conversion is complete")


root = tk.Tk()
root.title("Doc To Pdf Converter")
initialize_widgets()

root.mainloop()
