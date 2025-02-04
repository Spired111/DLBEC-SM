import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox
import pandas as pd
from tkinter import filedialog
from tkinter import Listbox, EXTENDED
import os
from PIL import Image, ImageTk

def load_excel_file():
    """Prompts the user to select a single Excel file and loads it."""
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        Messagebox.show_error("No file selected!", "Error")
        return None
    try:
        data = pd.read_excel(filepath, engine='openpyxl')
        return filepath, data
    except Exception as e:
        Messagebox.show_error(f"Error loading file {filepath}: {e}", "Error")
        return None

def omit_columns_from_file(filepath, data, columns_to_omit):
    """Omit selected columns from the loaded file."""
    data = data.drop(columns=columns_to_omit, errors='ignore')  # Drop columns
    save_omitted_file(data, filepath)

def save_omitted_file(data, original_filepath):
    """Saves the modified file with omitted columns to a new file."""
    # Create a 'Clean Files' folder if it doesn't exist
    directory = os.path.dirname(original_filepath)
    clean_folder = os.path.join(directory, "Clean Files")
    os.makedirs(clean_folder, exist_ok=True)
    
    # Generate the new filename by appending '_omitted_columns'
    filename = os.path.basename(original_filepath)
    new_filename = f"{os.path.splitext(filename)[0]}_omitted_columns.xlsx"
    save_path = os.path.join(clean_folder, new_filename)
    
    # Save the modified data to the new file
    try:
        data.to_excel(save_path, index=False, engine='openpyxl')
        Messagebox.show_info(f"File saved successfully: {save_path}", "Success")
    except Exception as e:
        Messagebox.show_error(f"Error saving file: {e}", "Error")

def main():
    app = ttk.Window(title="DLBEC Data Omitter", themename="darkly", size=(600, 450))
    
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
    filepath = None
    data = None
    columns_to_omit = []

    def on_choose_file():
        nonlocal filepath, data

        # Load the selected Excel file
        result = load_excel_file()
        if not result:
            return

        filepath, data = result

        # Get column names from the loaded file
        columns = data.columns.tolist()
        
        # Insert columns into the listbox
        column_listbox.delete(0, 'end')
        for col in columns:
            column_listbox.insert('end', col)

        # Enable the "Omit Columns and Save" button
        omit_button.config(state="normal")
    
    def on_omit_columns_and_save():
        nonlocal filepath, data, columns_to_omit
        
        if filepath is None or data is None:
            Messagebox.show_error("No file loaded!", "Error")
            return

        if not columns_to_omit:
            Messagebox.show_error("No columns selected to omit!", "Error")
            return

        # Omit the selected columns and save the new file
        omit_columns_from_file(filepath, data, columns_to_omit)

    def on_select_columns(event):
        """Handle the column selection event."""
        selected_columns = column_listbox.curselection()
        if selected_columns:
            # Clear previous selection
            columns_to_omit.clear()
            
            for col_idx in selected_columns:
                columns_to_omit.append(column_listbox.get(col_idx))
            
            if not columns_to_omit:
                Messagebox.show_warning("No columns selected!", "Warning")

    # Set up the button frame
    button_frame = ttk.Frame(app)
    button_frame.pack(pady=10)

    # Add a label describing the program
    program_label = ttk.Label(button_frame, text="Select a file and omit columns before saving.", bootstyle="info")
    program_label.pack(pady=4, side="top")
    
    # Set up the Choose File button
    choose_file_button = ttk.Button(button_frame, text="Choose file", command=on_choose_file, bootstyle="success")
    choose_file_button.pack(padx=5, side="bottom", expand=False)

    # Set up listbox frame for omitting columns
    omit_column_frame = ttk.Frame(app)
    omit_column_frame.pack(pady=10)

    # Label for the omit column feature
    omit_column_label = ttk.Label(omit_column_frame, text="Select columns to omit", bootstyle="info")
    omit_column_label.pack(pady=4, side="top")

    # Listbox for column selection (multi-select enabled)
    column_listbox = Listbox(omit_column_frame, selectmode=EXTENDED, height=8)
    column_listbox.pack(padx=5, pady=10, side="left", expand=True)
    column_listbox.bind('<<ListboxSelect>>', on_select_columns)

    # Set up save frame
    save_frame = ttk.Frame(app)
    save_frame.pack(pady=10)

    # Add a label describing how files are saved
    save_label = ttk.Label(save_frame, text="The file will be saved with omitted columns.", bootstyle="info")
    save_label.pack(pady=4, side="top")

    # Omit Columns and Save button
    omit_button = ttk.Button(save_frame, text="Omit Columns and Save", command=on_omit_columns_and_save, state="disabled", bootstyle="primary")
    omit_button.pack(pady=10, padx=10)

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
