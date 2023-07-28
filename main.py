import tkinter as tk
from tkinter import filedialog
import pandas as pd


def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file_paths = root.tk.splitlist(file_paths)

    for file_path in file_paths:
        # Read Excel file into a pandas DataFrame
        df = pd.read_excel(file_path)

        # Open a new window to input column headers and sorting orders
        sort_window = tk.Toplevel(root)
        sort_window.title("Sort Columns")

        def sort_and_save(data_frame):
            headers_and_orders = {}
            for i in range(len(header_entries)):
                header = header_entries[i].get()
                order = order_vars[i].get()
                if header and header in data_frame.columns:
                    headers_and_orders[header] = order

            # Sort the DataFrame based on the user-provided column headers and orders
            for header, ascending in headers_and_orders.items():
                data_frame = data_frame.sort_values(by=header, ascending=ascending)

            print(f"Sorted DataFrame based on selected columns:\n{data_frame}\n")

            # Save the sorted DataFrame back to the same Excel file
            data_frame.to_excel(file_path, index=False)
            print(f"Sorted data saved back to '{file_path}'.\n")

            # Close the sort window after sorting and saving
            sort_window.destroy()

        header_entries = []
        order_vars = []
        for column in df.columns:
            entry_frame = tk.Frame(sort_window)
            entry_frame.pack(pady=5)

            header_var = tk.StringVar()
            header_var.set(column)
            entry = tk.Entry(entry_frame, textvariable=header_var)
            entry.pack(side=tk.LEFT)
            header_entries.append(header_var)

            order_var = tk.BooleanVar()
            check_box = tk.Checkbutton(entry_frame, text="Ascending", variable=order_var)
            check_box.pack(side=tk.LEFT)
            order_vars.append(order_var)

        sort_button = tk.Button(sort_window, text="Sort and Save", command=lambda: sort_and_save(df))
        sort_button.pack(pady=5)


# Create the GUI
root = tk.Tk()
root.title("Excel to Pandas DataFrame")

button_select = tk.Button(root, text="Select Excel files", command=select_files)
button_select.pack(pady=10)

root.mainloop()
