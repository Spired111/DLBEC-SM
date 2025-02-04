import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox
import pandas as pd
from tkinter import filedialog
import os
from PIL import Image, ImageTk  # Importing Pillow for image handling

def load_excel_files():
    """Prompts the user to select multiple Excel files and loads them."""
    filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if not filepaths:
        Messagebox.show_error("No files selected!", "Error")
        return None
    files_data = []
    for filepath in filepaths:
        try:
            data = pd.read_excel(filepath, engine='openpyxl')
            files_data.append((filepath, data))
        except Exception as e:
            Messagebox.show_error(f"Error loading file {filepath}: {e}", "Error")
            continue
    return files_data

def get_unique_postcode_prefixes(df, column):
    """Extracts unique postcode prefixes (first two letters, case insensitive) from the specified column."""
    if column not in df.columns:
        raise ValueError(f"Selected column '{column}' not found in the dataset.")
    
    df[column] = df[column].astype(str)  # Ensure the column is treated as strings
    prefixes = df[column].apply(lambda x: x[:2].upper() if isinstance(x, str) else "").unique()
    return sorted(prefixes)

def filter_by_postcode(df, column, prefix):
    """Filters rows where the selected column contains the given postcode prefix (case insensitive)."""
    filtered_data = df[df[column].str.contains(prefix, case=False, na=False)]
    return filtered_data

def save_filtered_data(filtered_data, filepath, prefix):
    """Saves the filtered data to a new Excel file in the 'Clean Files' folder."""
    # Get the directory of the original file
    directory = os.path.dirname(filepath)
    
    # Ensure the "Clean Files" folder exists
    clean_folder = os.path.join(directory, "Clean Files")
    os.makedirs(clean_folder, exist_ok=True)
    
    # Generate the filename with one less row to account for headers
    num_rows = len(filtered_data) - 1  # Subtract 1 for the header
    filename = f"{prefix.upper()} Postcode ({num_rows}).xlsx" if num_rows >= 0 else f"{prefix.upper()} Postcode (0).xlsx"
    save_path = os.path.join(clean_folder, filename)
    
    # Save the filtered data to a new file
    try:
        filtered_data.to_excel(save_path, index=False, engine='openpyxl')
        Messagebox.show_info(f"Filtered data saved successfully to {save_path}", "Success")
    except Exception as e:
        Messagebox.show_error(f"Error saving file: {e}", "Error")

def main():
    app = ttk.Window(title="DLBEC SM", themename="darkly", size=(600, 400))
    
    # Set the window icon (make sure the .ico file is in the same directory)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(script_dir, "app_icon.ico")
    logo_path = os.path.join(script_dir, "logo.png")
    app.iconbitmap(icon_path)
    
    # Load logo image and display in the window
    try:
        logo_image = Image.open(logo_path)  # Replace with your logo filename
        logo_image = logo_image.resize((100, 100))
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = ttk.Label(app, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.pack(pady=10)
    except FileNotFoundError:
        Messagebox.show_warning("Logo image not found!", "Warning")

    # Create variables for data and selections
    selected_column = None
    selected_prefix = None
    files_data = None
    
    def on_choose_files():
        nonlocal files_data
        
        # Disable the button after selecting the files
        choose_file_button.config(state="disabled")

        # Load the selected Excel files
        files_data = load_excel_files()
        if not files_data:
            return
        
        # Enable and reset the combo boxes for column selection
        columns = list(files_data[0][1].columns)  # Get columns from the first file
        col_combo['values'] = columns
        col_combo.set("")  # Clear previous selection
        prefix_entry.delete(0, 'end')  # Clear the prefix entry
        col_combo.config(state="normal")
        save_button.config(state="disabled")  # Disable "Save Data" initially
    
    def on_column_select(event):
        nonlocal selected_column
        selected_column = col_combo.get()
        if selected_column:
            try:
                prefixes = get_unique_postcode_prefixes(files_data[0][1], selected_column)
            except ValueError as e:
                Messagebox.show_error(f"Error: {e}", "Error")
                return

    def on_save_data():
        nonlocal files_data, selected_column, selected_prefix
        
        # Get the prefix from the entry field
        selected_prefix = prefix_entry.get().strip()
        if not selected_prefix:
            Messagebox.show_error("Please enter a postcode prefix!", "Error")
            return

        for filepath, data in files_data:
            # Check if the column exists
            if selected_column not in data.columns:
                Messagebox.show_error(f"Column '{selected_column}' not found in file {filepath}!", "Error")
                continue
            
            # Filter and save the data
            filtered_data = filter_by_postcode(data, selected_column, selected_prefix)
            if filtered_data.empty:
                Messagebox.show_error(f"No matching entries found in file {filepath}!", "Error")
                continue
            
            save_filtered_data(filtered_data, filepath, selected_prefix)
    
    def on_prefix_entry_change(*args):
        """Enable Save Data button when the entry has text."""
        if prefix_entry.get().strip():  # If the entry box is not empty
            save_button.config(state="normal")
    
    # Set up the button frame
    button_frame = ttk.Frame(app)
    button_frame.pack(pady=10)

    # Add a label describing the program
    program_label = ttk.Label(button_frame, text="Choose one or more xlsx files to extract a specific postcode from them.", bootstyle="info")
    program_label.pack(pady=4, side="top")
    
    # Set up the Choose Files button
    choose_file_button = ttk.Button(button_frame, text="Choose files", command=on_choose_files, bootstyle="success")
    choose_file_button.pack(padx=5, side="bottom", expand=False)

    # Set up combo box frame for column selection
    combo_frame = ttk.Frame(app)
    combo_frame.pack(pady=10)

    # Add a label describing how to choose the columns
    choice_label = ttk.Label(combo_frame, text="Choose the column and enter the postcode prefix to filter by.", bootstyle="info")
    choice_label.pack(pady=4, side="top")

    # Column selection combo box
    col_combo = ttk.Combobox(combo_frame, state="disabled", bootstyle="dark")
    col_combo.bind("<<ComboboxSelected>>", on_column_select)
    col_combo.pack(padx=5, pady=10, side="left")

    # Postcode prefix entry box
    prefix_label = ttk.Label(combo_frame, text="Enter Postcode Prefix:", bootstyle="info")
    prefix_label.pack(pady=4, side="top")

    prefix_entry = ttk.Entry(combo_frame, bootstyle="dark")
    prefix_entry.bind("<KeyRelease>", on_prefix_entry_change)
    prefix_entry.pack(padx=5, pady=10, side="left")

    # Set up save frame
    save_frame = ttk.Frame(app)
    save_frame.pack(pady=10)

    # Add a label describing how files are saved
    save_label = ttk.Label(save_frame, text="Files will be saved to the 'Clean Files' folder.", bootstyle="info")
    save_label.pack(pady=4, side="top")

    # Save Data button
    save_button = ttk.Button(save_frame, text="Save Data", command=on_save_data, state="disabled", bootstyle="primary")
    save_button.pack(pady=10, padx=10)

    # Center the window on the screen
    window_width = 600
    window_height = 450
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    app.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    app.mainloop()

if __name__ == "__main__":
    main()
