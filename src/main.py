import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
from datetime import datetime
from converter import MarkdownToWordConverter

# Configure Advanced Logging in Main Script
log_file = "conversion_log.txt"
logging.basicConfig(filename=log_file, level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

# Define default folders
DEFAULT_TEMPLATE_FOLDER = r"C:\Users\jlawrence\OneDrive - Photronics\Documents\TemplateCity"
DEFAULT_MARKDOWN_FOLDER = r"C:\Users\jlawrence\OneDrive - Photronics\Documents\Markdown"
DEFAULT_OUTPUT_FOLDER = os.path.join(DEFAULT_MARKDOWN_FOLDER, "Output")

# Ensure the output directory exists
os.makedirs(DEFAULT_OUTPUT_FOLDER, exist_ok=True)

def log_message(message, level="info"):
    """Logs a message to the log file and prints it to the console."""
    print(message)
    if level == "debug":
        logging.debug(message)
    elif level == "info":
        logging.info(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "error":
        logging.error(message)

def get_file_path(title, file_types, initial_dir):
    """Open a file chooser dialog for selecting a file."""
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=file_types,
        initialdir=initial_dir  # Use default location
    )

    if file_path:
        log_message(f"Selected file: {file_path}", "debug")
    else:
        log_message(f"No file selected for {title}", "warning")

    return file_path if file_path else None

def get_save_location(base_name):
    """Automatically generate a timestamped output file."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_filename = f"{timestamp}_{base_name}.docx"
    output_path = os.path.join(DEFAULT_OUTPUT_FOLDER, output_filename)

    log_message(f"Generated timestamped output file: {output_path}", "debug")
    return output_path

def show_progress(step, total_steps, message, progress_bar, root):
    """Update the progress bar and display the current step."""
    progress_percentage = int((step / total_steps) * 100)
    progress_bar["value"] = progress_percentage
    root.update_idletasks()
    log_message(f"{progress_percentage}% - {message}", "debug")

def main():
    root = tk.Tk()
    root.title("Markdown to Word Converter")
    root.geometry("400x200")

    ttk.Label(root, text="Processing...").pack(pady=10)
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=20)
    root.update_idletasks()

    log_message("\n--- Markdown to Word Conversion Started ---", "debug")

    try:
        total_steps = 7  # Number of major steps

        # Step 1: Select Word Template
        show_progress(1, total_steps, "Selecting Word Template...", progress_bar, root)
        template_path = get_file_path(
            "Select Word Template",
            [("Word Templates", "*.docx"), ("All files", "*.*")],
            DEFAULT_TEMPLATE_FOLDER
        )
        if not template_path:
            log_message("No template selected. Exiting...", "warning")
            return

        # Step 2: Select Markdown File
        show_progress(2, total_steps, "Selecting Markdown File...", progress_bar, root)
        markdown_path = get_file_path(
            "Select Markdown File",
            [("Markdown files", "*.md"), ("Text files", "*.txt"), ("All files", "*.*")],
            DEFAULT_MARKDOWN_FOLDER
        )
        if not markdown_path:
            log_message("No markdown file selected. Exiting...", "warning")
            return

        # Step 3: Log the Start of Processing
        log_message(f"Processing Markdown: {markdown_path}", "debug")

        # Step 4: Ensure all required styles exist in the Word Template
        show_progress(4, total_steps, "Checking and updating template styles...", progress_bar, root)

        # Step 5: Generate timestamped output filename
        show_progress(5, total_steps, "Generating timestamped output filename...", progress_bar, root)
        base_name = os.path.splitext(os.path.basename(markdown_path))[0]
        output_path = get_save_location(base_name)

        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        os.makedirs(output_dir, exist_ok=True)

        # Step 6: Initialize and perform conversion
        show_progress(6, total_steps, "Converting Markdown to Word...", progress_bar, root)
        converter = MarkdownToWordConverter()
        output_file = converter.convert(template_path, markdown_path, output_dir)

        show_progress(7, total_steps, "Conversion complete!", progress_bar, root)

        log_message(f"Conversion completed successfully! Output file: {output_file}", "info")

        # Step 7: Ask user if they want to open the file
        if messagebox.askyesno("Success", "Would you like to open the converted document?"):
            subprocess.run(["start", output_file], shell=True)  # Open file in default Word application

    except Exception as e:
        error_message = f"Error during conversion: {str(e)}"
        log_message(error_message, "error")
        messagebox.showerror("Error", error_message)
        logging.error(error_message, exc_info=True)

    root.destroy()
    log_message("--- Markdown to Word Conversion Finished ---", "debug")

if __name__ == "__main__":
    main()