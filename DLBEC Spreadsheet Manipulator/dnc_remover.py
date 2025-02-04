import os
import pandas as pd
from tkinter import filedialog, Tk
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox

def load_excel_file():
    """Prompts the user to select an Excel file and loads it."""
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        Messagebox.show_error("No file selected!", "Error")
        return None
    try:
        data = pd.read_excel(filepath, engine='openpyxl')
        return filepath, data
    except Exception as e:
        Messagebox.show_error(f"Error loading file: {e}", "Error")
        return None, None

def get_phone_column(df, column_name):
    """Extracts and returns the phone number column from the DataFrame."""
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in the dataset.")
    return df[column_name].astype(str)

def remove_dnc_entries(main_data, dnc_data):
    """Removes entries from the main sheet based on DNC data's phone numbers."""
    dnc_phone_numbers = set(dnc_data["Telephone Number"].astype(str))
    filtered_data = main_data[~main_data["First Number"].astype(str).isin(dnc_phone_numbers)]
    return filtered_data

def main():
    app = Tk()
    app.title("DLBEC Objection Remover")
    app.geometry("600x400")
    app.iconbitmap("app_icon.ico")  # Set app icon

    # Apply the ttkbootstrap theme correctly
    app.tk_setPalette(background='#1E1E1E')  # Sets dark background
    style = ttk.Style("darkly")
    

    main_filepath = None
    main_data = None
    dnc_filepath = None
    dnc_data = None

    def load_main_sheet():
        nonlocal main_filepath, main_data
        main_filepath, main_data = load_excel_file()
        if main_data is not None:
            choose_dnc_button["state"] = "normal"  # Enable DNC button after main sheet is chosen

    def load_dnc_sheet():
        nonlocal dnc_filepath, dnc_data
        dnc_filepath, dnc_data = load_excel_file()
        if dnc_data is not None:
            # Process the removal of DNC entries
            try:
                filtered_data = remove_dnc_entries(main_data, dnc_data)
                
                # Calculate the number of entries removed
                entries_removed = len(main_data) - len(filtered_data)
                
                if entries_removed > 0:
                    # Generate the new filename
                    original_filename = os.path.basename(main_filepath)
                    prefix, total_entries = original_filename.rsplit("(", 1)
                    total_entries = int(total_entries.rstrip(").xlsx"))
                    new_filename = f"{prefix}({total_entries - entries_removed}).xlsx"
                    new_filepath = os.path.join(os.path.dirname(main_filepath), new_filename)
                    
                    # Save the filtered data to the new file
                    filtered_data.to_excel(new_filepath, index=False, engine='openpyxl')
                    
                    # Delete the original file
                    os.remove(main_filepath)
                    
                    Messagebox.show_info(f"{entries_removed} entries were removed.\nFiltered file saved as:\n{new_filename}", "Success")
                else:
                    # No entries were removed; do not delete the original file
                    Messagebox.show_info("No entries were removed.", "0 Entries Removed")
                
            except Exception as e:
                Messagebox.show_error(f"Error processing DNC data: {e}", "Error")

    # Buttons
    choose_main_button = ttk.Button(app, text="Choose Main Sheet", command=load_main_sheet, bootstyle="success")
    choose_main_button.pack(pady=10)
    
    choose_dnc_button = ttk.Button(app, text="Choose DNC Sheet", command=load_dnc_sheet, bootstyle="danger", state="disabled")
    choose_dnc_button.pack(pady=10)
    
    app.mainloop()

if __name__ == "__main__":
    main()
