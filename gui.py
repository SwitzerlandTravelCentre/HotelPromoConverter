from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox

from converter import transform_excel


def run_gui() -> None:
    root = tk.Tk()
    root.title("Hotel Data Power BI Converter")
    root.geometry("620x390")

    file_path_var = tk.StringVar()
    output_path_var = tk.StringVar()
    azure_maps_key_var = tk.StringVar()
    show_key_var = tk.BooleanVar(value=False)

    def choose_file() -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            file_path_var.set(path)

    def choose_output_dir() -> None:
        path = filedialog.askdirectory()
        if path:
            output_path_var.set(path)

    def convert_file() -> None:
        input_path = file_path_var.get()
        output_dir = output_path_var.get()
        azure_maps_key = azure_maps_key_var.get().strip() or None

        if not input_path or not output_dir:
            messagebox.showerror("Missing input", "Please select both a file and an output folder.")
            return

        try:
            result_path = transform_excel(
                input_path,
                output_dir,
                azure_maps_subscription_key=azure_maps_key,
            )
            messagebox.showinfo("Success", f"File saved as:\n{result_path}")
        except Exception as error:
            messagebox.showerror("Error", str(error))

    def toggle_key_visibility() -> None:
        azure_key_entry.configure(show="" if show_key_var.get() else "*")

    tk.Label(root, text="Step 1: Choose Swisshotels/STC Excel file (.xlsx)").pack(pady=5)
    tk.Entry(root, textvariable=file_path_var, width=78).pack()
    tk.Button(root, text="Browse", command=choose_file).pack(pady=5)

    tk.Label(root, text="Step 2: Choose output folder").pack(pady=5)
    tk.Entry(root, textvariable=output_path_var, width=78).pack()
    tk.Button(root, text="Browse", command=choose_output_dir).pack(pady=5)

    tk.Label(root, text="Step 3: Enter Azure Maps subscription key").pack(pady=5)
    azure_key_entry = tk.Entry(root, textvariable=azure_maps_key_var, width=78, show="*")
    azure_key_entry.pack()
    tk.Checkbutton(
        root,
        text="Show key",
        variable=show_key_var,
        command=toggle_key_visibility,
    ).pack(pady=5)

    tk.Button(
        root,
        text="Step 4: Convert and Save File",
        command=convert_file,
        bg="green",
        fg="white",
    ).pack(pady=15)

    tk.Label(
        root,
        text="Hotel Data Converter - anonymises and cleans hotel data exports for Power BI.",
        fg="gray",
    ).pack(pady=10)

    root.mainloop()
