from __future__ import annotations

import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from converter import transform_excel
from secure_store import (
    SecureStoreError,
    delete_azure_maps_key,
    has_saved_azure_maps_key,
    load_azure_maps_key,
    save_azure_maps_key,
)


def run_gui() -> None:
    root = tk.Tk()
    root.title("Hotel Data Power BI Converter")
    root.geometry("720x560")

    file_path_var = tk.StringVar()
    output_path_var = tk.StringVar()
    azure_maps_key_var = tk.StringVar()
    show_key_var = tk.BooleanVar(value=False)
    save_key_var = tk.BooleanVar(value=False)
    status_var = tk.StringVar(value="Ready")
    worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
    is_running = tk.BooleanVar(value=False)

    def choose_file() -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            file_path_var.set(path)

    def choose_output_dir() -> None:
        path = filedialog.askdirectory()
        if path:
            output_path_var.set(path)

    def append_log(message: str) -> None:
        log_text.configure(state="normal")
        log_text.insert("end", f"{message}\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    def set_running(running: bool) -> None:
        is_running.set(running)
        state = "disabled" if running else "normal"

        for widget in (
            browse_file_button,
            browse_output_button,
            azure_key_entry,
            show_key_check,
            save_key_check,
            forget_key_button,
            convert_button,
        ):
            widget.configure(state=state)

        if running:
            progress_bar.start(12)
        else:
            progress_bar.stop()

    def conversion_progress(message: str) -> None:
        worker_queue.put(("log", message))

    def run_conversion(input_path: str, output_dir: str, azure_maps_key: str | None, save_key: bool) -> None:
        try:
            if save_key and azure_maps_key:
                save_azure_maps_key(azure_maps_key)
                worker_queue.put(("log", "Saved Azure Maps key securely for this Windows user"))

            result_path = transform_excel(
                input_path,
                output_dir,
                azure_maps_subscription_key=azure_maps_key,
                progress_callback=conversion_progress,
            )
            worker_queue.put(("done", result_path))
        except Exception as error:
            worker_queue.put(("error", str(error)))

    def poll_worker_queue() -> None:
        try:
            while True:
                event, payload = worker_queue.get_nowait()

                if event == "log":
                    status_var.set(str(payload))
                    append_log(str(payload))
                elif event == "done":
                    set_running(False)
                    status_var.set("Done")
                    append_log(f"Saved file: {payload}")
                    messagebox.showinfo("Success", f"File saved as:\n{payload}")
                elif event == "error":
                    set_running(False)
                    status_var.set("Error")
                    append_log(f"ERROR: {payload}")
                    messagebox.showerror("Error", str(payload))
        except queue.Empty:
            pass

        root.after(100, poll_worker_queue)

    def convert_file() -> None:
        if is_running.get():
            return

        input_path = file_path_var.get()
        output_dir = output_path_var.get()
        azure_maps_key = azure_maps_key_var.get().strip() or None

        if not input_path or not output_dir:
            messagebox.showerror("Missing input", "Please select both a file and an output folder.")
            return

        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")
        append_log("Starting conversion")
        status_var.set("Starting conversion")
        set_running(True)

        threading.Thread(
            target=run_conversion,
            args=(input_path, output_dir, azure_maps_key, save_key_var.get()),
            daemon=True,
        ).start()

    def toggle_key_visibility() -> None:
        azure_key_entry.configure(show="" if show_key_var.get() else "*")

    def forget_saved_key() -> None:
        try:
            delete_azure_maps_key()
            save_key_var.set(False)
            append_log("Removed saved Azure Maps key")
            status_var.set("Saved key removed")
        except Exception as error:
            messagebox.showerror("Error", str(error))

    def load_saved_key_on_startup() -> None:
        try:
            if has_saved_azure_maps_key():
                azure_maps_key_var.set(load_azure_maps_key())
                save_key_var.set(True)
                append_log("Loaded saved Azure Maps key")
        except SecureStoreError as error:
            append_log(f"Could not load saved Azure Maps key: {error}")

    tk.Label(root, text="Step 1: Choose Swisshotels/STC Excel file (.xlsx)").pack(pady=5)
    tk.Entry(root, textvariable=file_path_var, width=78).pack()
    browse_file_button = tk.Button(root, text="Browse", command=choose_file)
    browse_file_button.pack(pady=5)

    tk.Label(root, text="Step 2: Choose output folder").pack(pady=5)
    tk.Entry(root, textvariable=output_path_var, width=78).pack()
    browse_output_button = tk.Button(root, text="Browse", command=choose_output_dir)
    browse_output_button.pack(pady=5)

    tk.Label(root, text="Step 3: Enter Azure Maps subscription key").pack(pady=5)
    azure_key_entry = tk.Entry(root, textvariable=azure_maps_key_var, width=78, show="*")
    azure_key_entry.pack()
    show_key_check = tk.Checkbutton(
        root,
        text="Show key",
        variable=show_key_var,
        command=toggle_key_visibility,
    )
    show_key_check.pack(pady=2)

    key_options = tk.Frame(root)
    key_options.pack(pady=2)
    save_key_check = tk.Checkbutton(
        key_options,
        text="Save key securely for next time",
        variable=save_key_var,
    )
    save_key_check.pack(side="left", padx=5)
    forget_key_button = tk.Button(key_options, text="Forget saved key", command=forget_saved_key)
    forget_key_button.pack(side="left", padx=5)

    convert_button = tk.Button(
        root,
        text="Step 4: Convert and Save File",
        command=convert_file,
        bg="green",
        fg="white",
    )
    convert_button.pack(pady=10)

    progress_bar = ttk.Progressbar(root, mode="indeterminate")
    progress_bar.pack(fill="x", padx=18, pady=4)

    tk.Label(root, textvariable=status_var, anchor="w").pack(fill="x", padx=18)

    log_text = scrolledtext.ScrolledText(root, height=9, width=86, state="disabled")
    log_text.pack(fill="both", expand=True, padx=18, pady=8)

    tk.Label(
        root,
        text="Hotel Data Converter - anonymises and cleans hotel data exports for Power BI.",
        fg="gray",
    ).pack(pady=6)

    load_saved_key_on_startup()
    root.after(100, poll_worker_queue)
    root.mainloop()
