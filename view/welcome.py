import tkinter as tk
from tkinter import filedialog
import os
import os.path
from docx2pdf import convert


class Welcome(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Conversor de doc a pdf")
        self.master.protocol("WM_DELETE_WINDOW", self.master.destroy)

        self.create_widgets()
        self.pack()
        self.destiny_folder = os.getcwd()
        self.files = []
        self.folders = []

    def create_widgets(self):
        self.select_folder_button = tk.Button(self, text="Seleccionar carpeta", command=self.add_folder)
        self.select_file_button = tk.Button(self, text="Seleccionar archivo", command=self.add_file)
        self.select_destiny_button = tk.Button(self, text="Seleccionar carpeta de destino", command=self.select_destiny_folder)
        self.convert_button = tk.Button(self, text="Convertir", command=self.convert_all)
        self.quit = tk.Button(self, text="Salir", fg="red",
                              command=self.master.destroy)
        self.file_list_label = tk.Label(self, text="Lista de archivos")
        self.file_list_box = tk.Listbox(self)
        self.folder_list_label = tk.Label(self, text="Lista de carpetas")
        self.folder_list_box = tk.Listbox(self)


        self.select_folder_button.grid(row = 1, column = 0, padx = 10, pady = 2)
        self.select_file_button.grid(row = 2, column = 0, padx = 10, pady = 2)
        self.select_destiny_button.grid(row = 3, column = 0, padx = 10, pady = 2)
        self.convert_button.grid(row = 4, column = 0, padx = 10, pady = 2)
        self.quit.grid(row = 5, column = 0, padx = 10, pady = 2)
        self.file_list_label.grid(row = 0, column = 1, padx = 10)
        self.file_list_box.grid(row = 1, column = 1, rowspan=4, padx = 10)
        self.folder_list_label.grid(row = 0, column = 2, padx = 10)
        self.folder_list_box.grid(row = 1, column = 2, rowspan=4, padx = 10)

    def add_folder(self):
        folder_path = filedialog.askdirectory(initialdir=os.getcwd(), title="Seleccionar carpeta")
        if(folder_path != ""):
            self.folders.append(folder_path)
            self.folder_list_box.insert(tk.END, folder_path.split("/")[-1])

    def add_file(self):
        file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Seleccionar archivo")
        if(file_path != ""):
            self.files.append(file_path)
            self.file_list_box.insert(tk.END, file_path.split("/")[-1])

    def select_destiny_folder(self):
        folder_path = filedialog.askdirectory(initialdir=os.getcwd(), title="Seleccionar carpeta")
        if(folder_path != ""):
            self.destiny_folder = folder_path

    def convert_all(self):
        for index, path in enumerate(self.files):
            convert(path)

        for index, folder_path in enumerate(self.folders):
            folder = os.listdir(folder_path)
            for i, filepath in enumerate(folder):
                if(os.path.isfile(filepath) and (filepath.endswith(".docx") or filepath.endswith(".doc"))):
                    convert(filepath)

        self.file_list_box.delete(0,tk.END)
        self.folder_list_box.delete(0,tk.END)

        self.files = []
        self.folders = []

root = tk.Tk()
app = Welcome(master=root)
app.mainloop()