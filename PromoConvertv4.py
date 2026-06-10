import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import random
import string


def clean_datetime_series(series):
    cleaned = series.replace(r'^\s*$', pd.NA, regex=True)
    return pd.to_datetime(cleaned, errors='coerce')


def transform_excel(input_path, output_dir):
    try:
        try:
            df = pd.read_excel(input_path)
        except ImportError as e:
            if "xlrd" in str(e):
                raise RuntimeError("This is a .xls file and this is not working. Please upload a .xlsx file. You can convert the file in Excel.")
            else:
                raise

        if 'Status' in df.columns:
            df['QtyCorrect'] = df['Status'].map({11: 1, 13: 2})
        else:
            raise RuntimeError("Column 'Status' not found in the Excel file.")

        if 'Rsv Number' in df.columns and 'Cancel Number' in df.columns:
            cancel_clean = df['Cancel Number'].astype('string').str.strip()
            for idx, row in df.iterrows():
                cancel_val = row['Cancel Number']
                if pd.notna(cancel_val) and str(cancel_val).strip() != "":
                    rsv_val = row['Rsv Number']
                    mask = (
                        (df['Rsv Number'] == rsv_val) &
                        (cancel_clean.isna() | (cancel_clean == "")) &
                        (df.index != idx)
                    )
                    df.loc[mask, 'Cancel Number'] = "cancelled"
        else:
            raise RuntimeError("Columns 'Rsv Number' and/or 'Cancel Number' not found in the Excel file.")

        if 'Rate' not in df.columns:
            raise RuntimeError("Column 'Rate' not found in the Excel file. Please check the exact header name.")

        df['NotTravelled'] = df.apply(
            lambda row: row['Rate'] if (
                pd.notna(row['Cancel Number']) and str(row['Cancel Number']).strip().lower() == "cancelled"
            ) else "",
            axis=1
        )

        df['travelled'] = df.apply(
            lambda row: row['Rate'] if (
                pd.isna(row['Cancel Number']) or str(row['Cancel Number']).strip() == ""
            ) else "",
            axis=1
        )

        if 'Rsv Date1' not in df.columns:
            raise RuntimeError("Column 'Rsv Date1' not found in the Excel file.")
        if 'Check In' not in df.columns or 'Check Out' not in df.columns:
            raise RuntimeError("Columns 'Check In' and/or 'Check Out' not found in the Excel file.")

        df['Rsv Date1'] = clean_datetime_series(df['Rsv Date1'])
        df['Check In'] = clean_datetime_series(df['Check In'])
        df['Check Out'] = clean_datetime_series(df['Check Out'])

        df['Rsv Date'] = df['Rsv Date1'].dt.date
        df['Rsv Time'] = df['Rsv Date1'].dt.time
        df['Check In'] = df['Check In'].dt.date
        df['Check Out'] = df['Check Out'].dt.date

        df.drop(columns=['Rsv Date1', 'Address2', 'Email', 'Remark'], inplace=True, errors='ignore')

        if 'Name' in df.columns:
            df['Name'] = df['Name'].astype(str).str[:3]
        if 'Firstname' in df.columns:
            df['Firstname'] = df['Firstname'].astype(str).str[:3]

        if 'Address1' in df.columns:
            unique_addresses = df['Address1'].dropna().unique()
            address_map = {
                addr: ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
                for addr in unique_addresses
            }
            df['Address1'] = df['Address1'].map(address_map)
            df.rename(columns={'Address1': 'AnonID'}, inplace=True)

        df['Booking Hour'] = df['Rsv Time'].apply(lambda t: t.hour if pd.notna(t) else pd.NA)

        # Add country suffix for Power BI maps
        for col in ['Hotel City', 'City']:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda x: f"{str(x).strip()}, Switzerland"
                    if pd.notna(x) and str(x).strip() != ""
                    else x
                )

        filename = "promodata.xlsx"
        full_output_path = os.path.join(output_dir, filename)
        df.to_excel(full_output_path, index=False, sheet_name="Sheet1")

        return full_output_path

    except Exception as e:
        raise RuntimeError(f"Error during transformation: {e}")


def run_gui():
    root = tk.Tk()
    root.title("CoopPromo Data Transformer")
    root.geometry("600x300")

    def choose_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        file_path_var.set(path)

    def choose_output_dir():
        path = filedialog.askdirectory()
        output_path_var.set(path)

    def convert_file():
        input_path = file_path_var.get()
        output_dir = output_path_var.get()
        if not input_path or not output_dir:
            messagebox.showerror("Missing Input", "Please select both a file and an output folder.")
            return
        try:
            result_path = transform_excel(input_path, output_dir)
            messagebox.showinfo("Success", f"File saved as:\n{result_path}")
        except Exception as err:
            messagebox.showerror("Error", str(err))

    tk.Label(root, text="Step 1: Choose Coop Promo Excel file").pack(pady=5)
    file_path_var = tk.StringVar()
    tk.Entry(root, textvariable=file_path_var, width=70).pack()
    tk.Button(root, text="Browse", command=choose_file).pack(pady=5)

    tk.Label(root, text="Step 2: Choose output folder").pack(pady=5)
    output_path_var = tk.StringVar()
    tk.Entry(root, textvariable=output_path_var, width=70).pack()
    tk.Button(root, text="Browse", command=choose_output_dir).pack(pady=5)

    tk.Button(root, text="Step 3: Convert and Save File", command=convert_file, bg="green", fg="white").pack(pady=20)

    tk.Label(root, text="CoopPromo Transformer – anonymises and cleans hoteldata export for analysis.", fg="gray").pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    run_gui()