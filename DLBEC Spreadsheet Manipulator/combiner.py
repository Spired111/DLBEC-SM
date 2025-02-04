import pandas as pd
import os
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog


# Declare files_data in the global scope
files_data = []


def load_excel_files():
    """Prompt the user to select multiple Excel files and load them."""
    global files_data
    filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if not filepaths:
        Messagebox.show_error("No files selected!", "Error")
        return []
    
    files_data = []
    for filepath in filepaths:
        try:
            data = pd.read_excel(filepath, engine='openpyxl')
            files_data.append((filepath, data))
        except Exception as e:
            Messagebox.show_error(f"Error loading file {filepath}: {e}", "Error")
            continue
    return files_data


def combine_data(files_data):
    """Combines all the data from the loaded files into a single DataFrame."""
    if not files_data:
        return None

    combined_data = pd.DataFrame()

    # Determine all unique columns from all files
    all_columns = set()
    for _, df in files_data:
        all_columns.update(df.columns)

    # Concatenate the files, ensuring columns match
    for _, df in files_data:
        # Add any missing columns to the dataframe
        missing_columns = all_columns - set(df.columns)
        for col in missing_columns:
            df[col] = None  # Add missing columns with NaN values

        # Reorder the columns to ensure consistency across all files
        df.columns = [str(col) for col in df.columns]  # Convert all column names to strings
        df = df[sorted(df.columns)]  # Sort the columns alphabetically

        # Append the data from this file to the combined DataFrame
        combined_data = pd.concat([combined_data, df], ignore_index=True, sort=False)

    return combined_data


def save_combined_data(combined_data):
    """Saves the combined data to a new Excel file."""
    if combined_data is None or combined_data.empty:
        Messagebox.show_error("No data to save.", "Error")
        return

    # Get the directory for saving the combined file
    directory = os.path.dirname(files_data[0][0])  # Directory of the first file
    save_folder = os.path.join(directory, "Combined Files")
    os.makedirs(save_folder, exist_ok=True)

    # Create the filename
    save_path = os.path.join(save_folder, "combined_data.xlsx")
    
    try:
        combined_data.to_excel(save_path, index=False, engine='openpyxl')
        Messagebox.show_info(f"Combined data saved successfully to {save_path}", "Success")
    except Exception as e:
        Messagebox.show_error(f"Error saving combined data: {e}", "Error")


def main():
    global files_data  # Ensure it's referenced as global in the main function
    app = ttk.Window(title="DLBEC Data Combiner", themename="darkly", size=(600, 400))
    
    # Load button
    def on_load_files():
        files_data = load_excel_files()  # No need for nonlocal, as files_data is now global
        if files_data:
            combine_button.config(state="normal")  # Enable the combine button after files are loaded
    
    # Combine and Save button
    def on_combine_and_save():
        combined_data = combine_data(files_data)
        save_combined_data(combined_data)

    # Set up the button frame
    button_frame = ttk.Frame(app)
    button_frame.pack(pady=10)

    # Add a label for program description
    program_label = ttk.Label(button_frame, text="Select multiple Excel files to combine into one.", bootstyle="info")
    program_label.pack(pady=4, side="top")

    # Set up the "Choose Files" button
    choose_file_button = ttk.Button(button_frame, text="Choose Files", command=on_load_files, bootstyle="success")
    choose_file_button.pack(padx=5, side="bottom", expand=False)

    # Combine and Save button
    combine_button = ttk.Button(button_frame, text="Combine and Save", command=on_combine_and_save, state="disabled", bootstyle="primary")
    combine_button.pack(pady=10, padx=10)

    # Center the window on the screen
    window_width = 600
    window_height = 400
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    app.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    app.mainloop()


if __name__ == "__main__":
    main()
