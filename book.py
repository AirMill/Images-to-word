import os
from PIL import Image
from docx import Document
from docx.shared import Inches
from tkinter import ttk, filedialog, messagebox
import tkinter as tk

# Function to get the creation date of a PNG image using EXIF data
def get_image_creation_date(image_path):
    try:
        img = Image.open(image_path)
        return img.getexif().get(36867, None)  # Exif DateTimeOriginal tag (for PNGs, this might not always be present)
    except:
        return None

# Function to resize image to fit one page in Word
def resize_image_to_fit_page(image_path, doc):
    img = Image.open(image_path)
    width, height = img.size

    # Word document page dimensions (letter size in inches: 8.5 x 11)
    max_width_in_inches = 8.0  # Leave some margin for better fit
    max_height_in_inches = 10.5

    # Convert inches to pixels (assuming 96 dpi)
    max_width_px = max_width_in_inches * 96
    max_height_px = max_height_in_inches * 96

    # Resize the image while maintaining aspect ratio
    aspect_ratio = width / height
    if width > height:  # Landscape mode
        new_width = min(max_width_px, width)
        new_height = new_width / aspect_ratio
    else:  # Portrait mode or square
        new_height = min(max_height_px, height)
        new_width = new_height * aspect_ratio

    # Add resized image to the document
    doc.add_picture(image_path, width=Inches(new_width / 96), height=Inches(new_height / 96))

# Function to create a Word document from PNG images in the selected folder, with a progress bar
def create_word_from_images(image_folder, progress_var, progress_bar, root):
    # Create a new Word document
    doc = Document()
    output_doc = os.path.join(image_folder, 'png_image_document.docx')

    # Get list of all PNG files in the folder
    image_files = [f for f in os.listdir(image_folder) if f.lower().endswith('.png')]

    # Sort PNG images by creation date
    image_files.sort(key=lambda f: get_image_creation_date(os.path.join(image_folder, f)) or '')

    # Initialize the progress bar
    total_files = len(image_files)
    progress_step = 100 / total_files if total_files > 0 else 1
    progress_var.set(0)
    progress_bar.update()

    # Add sorted PNG images to the Word document and update progress bar
    for idx, image_file in enumerate(image_files):
        img_path = os.path.join(image_folder, image_file)
        resize_image_to_fit_page(img_path, doc)

        # Update the progress bar
        progress_var.set((idx + 1) * progress_step)
        progress_bar.update()

    # Save the document in the same folder
    doc.save(output_doc)

    # Notify the user and reset the progress bar
    messagebox.showinfo("Success", f"Document saved as: {output_doc}")
    progress_var.set(0)
    progress_bar.update()

# Function to open folder selection dialog and process images
def select_folder(progress_var, progress_bar, root):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        create_word_from_images(folder_selected, progress_var, progress_bar, root)

# Create the main GUI window with a progress bar
def main():
    root = tk.Tk()
    root.title("PNG to Word Document")

    # Set window size
    root.geometry('400x200')

    # Progress bar variable
    progress_var = tk.DoubleVar()

    # Add a button to select the folder
    select_button = tk.Button(root, text="Select Folder", command=lambda: select_folder(progress_var, progress_bar, root), padx=20, pady=10)
    select_button.pack(pady=20)

    # Add a progress bar widget
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(fill=tk.X, padx=20, pady=10)

    # Start the GUI loop
    root.mainloop()

# Run the main function to start the GUI
if __name__ == "__main__":
    main()
