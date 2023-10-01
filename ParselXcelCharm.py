import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import threading
import sys

# Use the code snippet to create the image path
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

image_path = os.path.join(script_dir, "Nature.png")
temp_code_path = os.path.join(script_dir, "temp_code.py")

def execute_code():
    global temp_code_path

    selected_code = var.get()

    code_paths = {
        1: "Germany.py",
        2: "Brazil.py",
        3: "Evobus.py",
        4: "India.py"
    }
    code_path = code_paths.get(selected_code)

    pdf_path = pdf_path_entry.get()

    # Read the content of the code file
    with open(code_path, 'r') as f:
        code_content = f.read()

    # Replace the OLD_PDF_PATH placeholder with the actual PDF path
    updated_code = code_content.replace("OLD_PDF_PATH = r\"NEW_PDF_PATH\"", f"OLD_PDF_PATH = {repr(pdf_path)}")

    # Write the updated code to the temporary file
    with open(temp_code_path, 'w') as f:
        f.write(updated_code)

    def run_code():
        try:
            subprocess.run(['python', temp_code_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
            messagebox.showinfo("Execution Complete", "Code execution is done.")
        except Exception as e:
            messagebox.showerror("Error", f"Code execution failed: {str(e)}")
        finally:
            # Clean up the temporary code file
            os.remove(temp_code_path)

    # Run the code execution in a separate thread
    threading.Thread(target=run_code).start()

def browse_pdf():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, pdf_file)

# Create and configure the main window
root = tk.Tk()
root.title("PDF Processing Interface")

background_image = tk.PhotoImage(file=image_path)
background_label = tk.Label(root, image=background_image)
background_label.place(relwidth=1, relheight=1)

pdf_path_label = tk.Label(root, text="PDF File Path:")
pdf_path_label.grid(row=0, column=0, padx=10, pady=10)

pdf_path_entry = tk.Entry(root, width=50)
pdf_path_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="Browse", command=browse_pdf)
browse_button.grid(row=0, column=2, padx=10, pady=10)

var = tk.IntVar()

# Create a frame for radio buttons
radio_frame = tk.Frame(root)
radio_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

germany_radio = tk.Radiobutton(radio_frame, text="Germany", variable=var, value=1)
germany_radio.grid(row=0, column=0, padx=5)

brazil_radio = tk.Radiobutton(radio_frame, text="Brazil", variable=var, value=2)
brazil_radio.grid(row=0, column=1, padx=5)

evobus_radio = tk.Radiobutton(radio_frame, text="Evobus", variable=var, value=3)
evobus_radio.grid(row=0, column=2, padx=5)

india_radio = tk.Radiobutton(radio_frame, text="India", variable=var, value=4)
india_radio.grid(row=0, column=3, padx=5)

execute_button = tk.Button(root, text="Execute", command=execute_code)
execute_button.grid(row=2, column=0, columnspan=3, padx=10, pady=20)

# Center the window on the screen
root.geometry("600x300+300+200")

# Allow resizing
root.resizable(True, True)

# Start the main event loop
root.mainloop()
