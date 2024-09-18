import os
import threading
from tkinter import Tk, Label, Button, filedialog, messagebox
from tkinter import ttk  # For the progress bar
from docx2pdf import convert

# Create a global variable to control the cancellation
cancel_event = threading.Event()
conversion_thread = None  # Initialize conversion_thread as None

# Function to select a folder containing .docx files
def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        # Move to the next screen for confirmation
        show_confirmation_screen(folder_path, is_folder=True)

# Function to select one or more files
def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Word Files", "*.docx")])
    if file_paths:
        # Move to the next screen for confirmation
        show_confirmation_screen(file_paths, is_folder=False)

# Function to confirm and start the conversion process (either files or folder)
def show_confirmation_screen(path, is_folder):
    # Clear the screen
    for widget in root.winfo_children():
        widget.destroy()

    # Reset cancel event
    cancel_event.clear()

    if is_folder:
        Label(root, text=f"Selected Folder:\n{path}", wraplength=500, padx=20, pady=10).pack()
    else:
        Label(root, text=f"Selected Files:\n{', '.join(path)}", wraplength=500, padx=20, pady=10).pack()

    # Create a progress bar
    global progress_bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
    progress_bar.pack(pady=20)

    # Frame to hold buttons
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)

    # Button to start conversion
    start_button = Button(button_frame, text="Start Conversion", command=lambda: run_conversion_in_thread(path, is_folder), padx=20, pady=10)
    start_button.grid(row=0, column=0, padx=10)

    # Button to cancel the conversion
    cancel_button = Button(button_frame, text="Cancel", command=cancel_conversion, padx=20, pady=10)
    cancel_button.grid(row=0, column=1, padx=10)

# Function to cancel the conversion process
def cancel_conversion():
    global conversion_thread
    if conversion_thread is not None and conversion_thread.is_alive():
        cancel_event.set()  # Trigger the cancel event
    else:
        # If no conversion is running, simply navigate back
        show_initial_screen()

# Function to convert all .docx files to PDF with progress update
def start_conversion(path, is_folder):
    if is_folder:
        # Get all .docx files from the selected folder
        docx_files = [f for f in os.listdir(path) if f.endswith('.docx')]
        file_paths = [os.path.join(path, f) for f in docx_files]
    else:
        # Use the selected files
        file_paths = list(path)

    if not file_paths:
        messagebox.showerror("Error", "No .docx files found.")
        return

    # Set up progress bar parameters
    progress_bar["maximum"] = len(file_paths)
    progress_bar["value"] = 0

    # Convert each .docx file to PDF
    for index, docx_file in enumerate(file_paths):
        if cancel_event.is_set():
            # Show cancellation message on the main thread
            root.after(0, lambda: messagebox.showwarning("Cancelled", "Conversion process has been cancelled."))
            # Navigate back to the initial screen on the main thread
            root.after(0, lambda: show_initial_screen())
            break

        try:
            convert(docx_file)
        except Exception as e:
            # Handle individual file conversion errors without stopping the entire process
            root.after(0, lambda: messagebox.showerror("Error", f"Failed to convert {os.path.basename(docx_file)}.\nError: {str(e)}"))
            continue

        # Update progress bar on the main thread
        root.after(0, lambda val=index+1: progress_bar.configure(value=val))

    else:
        # If the loop wasn't broken, show success message
        root.after(0, lambda: messagebox.showinfo("Success", "All .docx files have been converted to PDF!"))
        # Navigate back to the initial screen on the main thread
        root.after(0, lambda: show_initial_screen())

# Function to run the conversion process in a separate thread
def run_conversion_in_thread(path, is_folder):
    global conversion_thread
    conversion_thread = threading.Thread(target=start_conversion, args=(path, is_folder), daemon=True)
    conversion_thread.start()

# Function to show the initial screen (Folder/File Selection)
def show_initial_screen():
    # Clear the screen
    for widget in root.winfo_children():
        widget.destroy()

    # Screen 1: Select folder or files
    Label(root, text="Welcome to Word to PDF Converter!", font=("Arial", 18), padx=20, pady=20).pack()
    Label(root, text="Select a folder or files containing .docx to convert them to PDF.", font=("Arial", 12), padx=20, pady=10).pack()

    # Frame to hold buttons
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=20)

    # Button to select folder
    Button(button_frame, text="Select Folder", command=select_folder, font=("Arial", 12), padx=20, pady=10).grid(row=0, column=0, padx=10)

    # Button to select files
    Button(button_frame, text="Select Files", command=select_files, font=("Arial", 12), padx=20, pady=10).grid(row=0, column=1, padx=10)

# Initialize tkinter window
root = Tk()
root.title("Word to PDF Converter")

# Set window size and make it resizable
root.geometry("600x400")  # Increased size
root.resizable(False, False)  # Prevent resizing for consistent layout

# Center the window on the screen
root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

# Show the initial screen
show_initial_screen()

# Start the tkinter event loop
root.mainloop()
