import tkinter as tk
from tkinter import filedialog

import processor


class GUI:
    def __init__(self, root):
        self.file_path = None
        self.root = root
        self.root.title("File Selection Dialog")
        self.root.geometry("160x120")

        # Create a button to open the file dialog
        self.open_button = tk.Button(self.root, text="Open File", command=self.open_file_dialog)
        self.open_button.pack(pady=10)
        # Create a label to display the selected file path
        self.file_path_label = tk.Label(self.root, text="No file selected.")
        self.file_path_label.pack()
        # Create a checkbox to toggle detailed view
        self.checkbox_var = tk.IntVar()
        self.detailed_cb = tk.Checkbutton(self.root, text="Detailed", variable=self.checkbox_var)
        self.detailed_cb.pack()
        # Create run button
        self.run_button = tk.Button(self.root, text="Save", command=self.run)
        self.run_button.pack()

    def open_file_dialog(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            self.file_path_label.config(text="Selected File: " + self.file_path)
        else:
            self.file_path_label.config(text="No file selected.")

    def handle_ready(self):
        self.file_path_label.config(text="Готово!")

    def run(self):
        processor.process_file(src_xml=self.file_path, detailed=self.checkbox_var.get(), callback=self.handle_ready())


def show_gui():
    # Create the main application window
    window = tk.Tk()
    gui = GUI(window)
    # Run the main event loop
    window.mainloop()
