from docx2pdf import convert
import os
from tkinter import filedialog
import tkinter as tk

root = tk.Tk()
root.title("Doc To Pdf Converter")

canvas1 = tk.Canvas(root, width=400, height=300, relief='raised')
canvas1.pack()

label1 = tk.Label(root, text='Doc to Pdf Converter')
label1.config(font=('helvetica', 14))
canvas1.create_window(200, 45, window=label1)

label2 = tk.Label(root, text='Choose Your Docx File :')
label2.config(font=('helvetica', 10))
canvas1.create_window(200, 100, window=label2)


def path_name():
    paths = filedialog.askopenfilenames(title="Select Docx Files")
    button2 = tk.Button(text='Convert', command=lambda: convert_files(paths), bg='brown', fg='white',
                        font=('helvetica', 9, 'bold'))
    canvas1.create_window(200, 210, window=button2)

    def convert_files(paths):
        for path in paths:
            output_filename = os.path.splitext(os.path.basename(path))[0] + ".pdf"
            convert(path, output_filename)
            label5 = tk.Label(root, text=f"Converted {output_filename}", font=('helvetica', 10))
            canvas1.create_window(200, 260 + (20 * paths.index(path)), window=label5)
        label5 = tk.Label(root, text="All conversions completed!", font=('helvetica', 10))
        canvas1.create_window(200, 260 + (20 * (len(paths) - 1)), window=label5)


button1 = tk.Button(text='Browse Your File', command=path_name, bg='grey', fg='black', font=('helvetica', 9, 'bold'))
canvas1.create_window(200, 150, window=button1)

root.mainloop()
