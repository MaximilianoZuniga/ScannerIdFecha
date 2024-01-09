import tkinter as tk
from tkinter import ttk,messagebox
import keyboard
from datetime import datetime
import openpyxl
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Lector de Barras App")
        self.root.geometry("800x600")

        self.frame_buttons = ttk.Frame(root)
        self.frame_buttons.pack(side=tk.TOP, fill=tk.X)

        self.start_button = tk.Button(self.frame_buttons, text="Iniciar Lectura", command=self.start_reading)
        self.start_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.stop_button = tk.Button(self.frame_buttons, text="Detener Lectura", command=self.stop_reading, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.export_button = tk.Button(self.frame_buttons, text="Exportar a Excel", command=self.export_to_excel, state="disabled")
        self.export_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.frame_data = ttk.Frame(root)
        self.frame_data.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.frame_data, columns=("Folio de rendición", "Fecha"))
        self.tree.heading("Folio de rendición", text="Código de Barras")
        self.tree.heading("Fecha", text="Fecha")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = ttk.Scrollbar(self.frame_data, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.current_barcode = ""
        self.data_list = []

    def start_reading(self):
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.export_button.config(state="disabled")
        keyboard.hook(self.key_callback)

    def stop_reading(self):
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.export_button.config(state="normal")
        keyboard.unhook_all()

    def key_callback(self, event):
        if event.event_type == keyboard.KEY_DOWN:
            if event.name.isnumeric() or event.name == "enter":
                self.current_barcode += event.name if event.name.isnumeric() else ""

            if event.name == "enter" and self.current_barcode:
                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                data_entry = (self.current_barcode, current_time)
                self.data_list.append(data_entry)
                self.update_data_list()
                self.current_barcode = ""

    def update_data_list(self):
        self.tree.delete(*self.tree.get_children())
        for data_entry in self.data_list:
            self.tree.insert("", "end", values=data_entry)

    def export_to_excel(self):
        if self.data_list:
            base_file_name = "data_export.xlsx"
            file_path = base_file_name

            # Check if the file already exists
            counter = 1
            while os.path.exists(file_path):
                counter += 1
                file_path = f"data_export_{counter}.xlsx"

            workbook = openpyxl.Workbook()
            sheet = workbook.active

            # Add headers
            sheet.append(["Código de Barras", "Fecha"])

            # Add data
            for data_entry in self.data_list:
                sheet.append(data_entry)

            # Save the workbook
            workbook.save(file_path)
            self.export_button.config(state="disabled")
            messagebox.showinfo("Exportación Exitosa", f"Los datos se han exportado correctamente a {file_path}")

            # Clear the data in the Treeview
            self.data_list = []
            self.update_data_list()

    def on_closing(self):
        self.root.destroy()
        keyboard.unhook_all()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()