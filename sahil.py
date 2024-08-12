import os
import platform
import logging
import csv
import textwrap
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry  # For date picker
from PIL import Image, ImageDraw, ImageFont, ImageTk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch, mm
import fitz  # PyMuPDF
import datetime
logging.basicConfig(filename='app.log', level=logging.ERROR, format='%(asctime)s - %(message)s')
# Update these paths according to your file locations
template_path = r'C:\Users\Shree\Desktop\idcard\projectidcard\ritika\ST.png'
image_folder = r'C:\Users\Shree\Desktop\idcard\projectidcard\ritika\downloaded_images'
qr_folder = r'C:\Users\Shree\Desktop\idcard\projectidcard\ritika\ST_output_qr_codes'
output_folder = r'C:\Users\Shree\Desktop\output'
csv_data = None
treeview = None
checkbox_vars = []

def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        display_pdf(pdf_path)
        try:
            if platform.system() == "Windows":
                os.startfile(pdf_path, "open")
        except Exception as e:
            messagebox.showerror("Error", f"Unable to open PDF file: {str(e)}")

def display_pdf(pdf_path):
    for widget in pdf_frame.winfo_children():
        widget.destroy()
    
    try:
        doc = fitz.open(pdf_path)
        num_pages = doc.page_count

        # Add a scrollbar
        scrollbar = tk.Scrollbar(pdf_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create a canvas to display the PDF pages
        pdf_canvas = tk.Canvas(pdf_frame, width=800, height=600, yscrollcommand=scrollbar.set)
        pdf_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure the scrollbar
        scrollbar.config(command=pdf_canvas.yview)

        # Create a frame to hold the PDF pages inside the canvas
        inner_frame = tk.Frame(pdf_canvas)
        pdf_canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

        # Display each page of the PDF
        for i in range(num_pages):
            page = doc.load_page(i)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_tk = ImageTk.PhotoImage(img)
            page_label = tk.Label(inner_frame, image=img_tk)
            page_label.image = img_tk  # Keep a reference to avoid garbage collection
            page_label.pack(fill=tk.X)

        # Set the scroll region after creating all the pages
        inner_frame.update_idletasks()
        pdf_canvas.config(scrollregion=pdf_canvas.bbox(tk.ALL))

    except Exception as e:
        messagebox.showerror("Error", f"Error displaying PDF: {str(e)}")

def resize_canvas(canvas):
    # Update the canvas size based on the frame size
    canvas.config(width=pdf_frame.winfo_width(), height=pdf_frame.winfo_height())
    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox(tk.ALL))

def preprocess_image(image_path):
    try:
        input_image = Image.open(image_path)
        final_image = input_image.convert("RGB")
        return final_image
    except Exception as e:
        messagebox.showerror("Error", f"Error opening image at {image_path}: {str(e)}")
        return None

def create_id_cards(template_path, image_folder, qr_folder, selected_indices):
    if csv_data is None or csv_data.empty:
        messagebox.showwarning("Warning", "Please load a CSV file first.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    rows_to_generate = csv_data.iloc[selected_indices]

    images = []
    for _, row in rows_to_generate.iterrows():
        card = generate_card(row, template_path, image_folder, qr_folder)
        if card:
            # Ensure the image is in RGB mode before saving as JPEG
            if card.mode in ("RGBA", "P"):
                card = card.convert("RGB")
                
            image_path = os.path.join(output_folder, f"{row['ID']}.jpg")
            try:
                card.save(image_path, "JPEG")
                images.append(image_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save image for ID {row['ID']}: {e}")

    if images:
        output_pdf_path = os.path.join(output_folder, "output.pdf")
        pdf_path = create_pdf(images, output_pdf_path)
        if pdf_path:
            messagebox.showinfo("Info", f"PDF created successfully: {pdf_path}")
            browse_pdf()
        else:
            messagebox.showerror("Error", "Failed to create PDF.")
    else:
        messagebox.showwarning("Warning", "No valid images to create ID cards.")

def generate_card(data, template_path, image_folder, qr_folder):
    pic_id = str(data.get('ID', ''))
    if not pic_id:
        messagebox.showwarning("Warning", f"Skipping record with missing ID: {data}")
        return None

    pic_path = os.path.join(image_folder, f"{pic_id}.jpg")
    qr_path = os.path.join(qr_folder, f"{pic_id}.png")

    if not os.path.exists(pic_path):
        messagebox.showerror("Error", f"Image not found for ID: {pic_id} at path: {pic_path}")
        return None

    if not os.path.exists(qr_path):
        messagebox.showerror("Error", f"QR code not found for ID: {pic_id} at path: {qr_path}")
        return None

    preprocessed_pic = preprocess_image(pic_path)
    if preprocessed_pic is None:
        return None

    try:
        preprocessed_pic = preprocessed_pic.resize((144, 145))
        template = Image.open(template_path)
        qr = Image.open(qr_path).resize((161, 159))

        template.paste(preprocessed_pic, (27, 113, 171, 258))
        template.paste(qr, (497, 109, 658, 268))

        draw = ImageDraw.Draw(template)

        try:
            font_path = "C:\\WINDOWS\\FONTS\\ARIAL.TTF"
            name_font = ImageFont.truetype(font_path, size=18)
        except IOError:
            name_font = ImageFont.load_default()

        wrapped_div = textwrap.fill(str(data.get('Division/Section', '')), width=22).title()
        draw.text((311, 121), wrapped_div, font=name_font, fill='black')

        division_input = data.get('Division/Section', '')
        head_name = get_head_by_division(division_input)
        wrapped_supri = textwrap.fill(str(head_name), width=20).title()
        draw.text((311, 170), wrapped_supri, font=name_font, fill='black')

        university = data.get('University', 'Not Available')
        draw.text((200, 356), university, font=name_font, fill='black')

        draw.text((305, 219), data.get('Internship Start Date', ''), font=name_font, fill='black')
        draw.text((303, 266), data.get('Internship End Date', ''), font=name_font, fill='black')
        draw.text((300, 312), str(data.get('Mobile', '')), font=name_font, fill='black')
        draw.text((621, 283), str(data.get('ID', '')), font=name_font, fill='black')

        wrapped_name = center_align_text_wrapper(data.get('Name', ''), width=22)
        name_bbox = name_font.getbbox(wrapped_name)
        name_width = name_bbox[2] - name_bbox[0]
        center_x = ((198 - name_width) / 2)
        draw.text((center_x, 260), wrapped_name, font=name_font, fill='black')

        return template

    except Exception as e:
        messagebox.showerror("Error", f"Error generating card for ID: {pic_id}. Error: {str(e)}")
        return None


def center_align_text_wrapper(text, width=15):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        if len(current_line) + len(word) + 1 <= width:
            current_line += word + " "
        else:
            lines.append(current_line[:-1])
            current_line = word + " "

    lines.append(current_line[:-1])
    centered_lines = [line.center(width) for line in lines]
    centered_text = "\n".join(centered_lines)
    return centered_text

def get_head_by_division(division_name):
    divisions = {
        "Advanced Information Technologies Group": "Dr. Sanjay Singh",
        "Societal Electronics Group": "Dr. Udit Narayan Pal",
        "Industrial Automation": "Dr. S.S. Sadistap",
        "Vacuum Electronic Devices Group": "Dr. Sanjay Kr. Ghosh",
        "High-Frequency Devices & System Group": "Dr. Ayan Bandhopadhyay",
        "Semiconductor Sensors & Microsystems Group": "Dr. Suchandan Pal",
        "Semiconductor Process Technology Group": "Dr. Kuldip Singh",
        "Industrial R & D": "Mr. Ashok Chauhan",
        "High Power Microwave Systems Group": "Dr. Anirban Bera",
    }

    division_name = division_name.strip().title()
    return divisions.get(division_name, "Division not found or head information not available.")

def create_pdf(images, pdf_path):
    try:
        c = canvas.Canvas(pdf_path, pagesize=letter)

        grid_width = 2
        grid_height = 4
        image_width = 3.575 * inch
        image_height = 2.325 * inch
        spacing_x = 1.5 * mm
        spacing_y = 1.5 * mm

        total_width = grid_width * (image_width + spacing_x)
        total_height = grid_height * (image_height + spacing_y)

        current_page = 0

        for i, image in enumerate(images):
            col = i % grid_width
            row = i // grid_width

            if i > 0 and i % (grid_width * grid_height) == 0:
                current_page += 1
                c.showPage()

            start_x = (letter[0] - total_width) / 2
            start_y = (letter[1] - total_height) / 2 - current_page * total_height

            x = start_x + col * (image_width + spacing_x)
            y = start_y + row * (image_height + spacing_y)

            c.drawInlineImage(image, x, y, width=image_width, height=image_height)

        c.save()
        return pdf_path

    except Exception as e:
        logging.error(f"Error creating PDF: {str(e)}")
        return None
# Function to select and load a CSV file
def select_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        try:
            global csv_data
            csv_data = pd.read_csv(file_path)
            display_csv_data(csv_data)
            messagebox.showinfo("Info", "CSV file loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Error reading CSV file: {str(e)}")
def display_csv_data(data):
    global treeview, checkbox_vars

    # Clear any existing widgets in the treeview_frame
    for widget in treeview_frame.winfo_children():
        widget.destroy()

    # Create a canvas to enable scrolling for both Treeview and checkboxes
    canvas = tk.Canvas(treeview_frame)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add a scrollbar to the canvas
    scrollbar = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.LEFT, fill=tk.Y)

    # Configure the canvas to work with the scrollbar
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Create a frame inside the canvas to hold the Treeview and checkboxes
    content_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=content_frame, anchor="nw")

    # Create a Treeview widget within the content frame
    treeview = ttk.Treeview(content_frame, columns=list(data.columns), show="headings")
    treeview.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    for col in data.columns:
        treeview.heading(col, text=col)
        treeview.column(col, width=100)

    # Insert the CSV data into the Treeview
    for i, row in data.iterrows():
        treeview.insert("", "end", values=list(row))

    # Create a frame for the checkboxes aligned with the Treeview rows
    checkbox_frame = tk.Frame(content_frame)
    checkbox_frame.pack(side=tk.LEFT, fill=tk.Y)

    # Create checkboxes aligned with each ID
    checkbox_vars = []
    for i in range(len(data)):
        var = tk.BooleanVar()
        chk = tk.Checkbutton(checkbox_frame, variable=var)
        chk.pack(anchor='w', pady=1.5)  # Adjust padding as necessary to align with rows
        checkbox_vars.append(var)

    # Update the scroll region of the canvas to match the content frame
    content_frame.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

def make_editable(treeview, selected_item):
    if selected_item is None:
        messagebox.showwarning("Warning", "No item selected for editing.")
        return
    
    values = treeview.item(selected_item, "values")
    index = treeview.index(selected_item)

    # Create a new window for editing
    edit_window = tk.Toplevel(root)
    edit_window.title("Edit Row")

    # Create entries for each column
    entries = []
    for i, col in enumerate(treeview["columns"]):
        tk.Label(edit_window, text=col).grid(row=i, column=0)
        
        if "Date" in col:  # If the column is a date, use a DateEntry widget
            # Try to parse the date in the expected format
            try:
                date_value = datetime.datetime.strptime(values[i], '%m/%d/%Y').date()  # Adjust format as needed
            except ValueError:
                date_value = datetime.datetime.strptime(values[i], '%Y-%m-%d').date()  # Adjust format as needed

            entry = DateEntry(edit_window, date_pattern='y-mm-dd')
            entry.set_date(date_value)  # Pre-fill with existing date
        else:
            entry = tk.Entry(edit_window)
            entry.insert(0, values[i])
        
        entry.grid(row=i, column=1)
        entries.append(entry)

    # Save changes button
    def save_changes():
        new_values = [entry.get() if not isinstance(entry, DateEntry) else entry.get_date().strftime('%Y-%m-%d') for entry in entries]
        csv_data.loc[index, :] = new_values
        treeview.item(selected_item, values=new_values)
        edit_window.destroy()

    tk.Button(edit_window, text="Save Changes", command=save_changes).grid(row=len(treeview["columns"]), column=0, columnspan=2)
def get_selected_rows_indices():
    global checkbox_vars
    selected_indices = [i for i, var in enumerate(checkbox_vars) if var.get()]
    return selected_indices


def on_generate_id_cards():
    selected_indices = get_selected_rows_indices()
    if not selected_indices:
        messagebox.showwarning("Warning", "Please select at least one row to generate ID cards.")
        return
    create_id_cards(template_path, image_folder, qr_folder, selected_indices=selected_indices)

def create_id_cards(template_path, image_folder, qr_folder, generation_type=None, selected_indices=None):
    if csv_data is None or csv_data.empty:
        messagebox.showwarning("Warning", "Please load a CSV file first.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    if selected_indices:
        rows_to_generate = csv_data.iloc[selected_indices]
    else:
        rows_to_generate = csv_data

    images = []
    for _, row in rows_to_generate.iterrows():
        card = generate_card(row, template_path, image_folder, qr_folder)
        if card:
            # Ensure the image is in RGB mode before saving as JPEG
            if card.mode in ("RGBA", "P"):
                card = card.convert("RGB")
                
            image_path = os.path.join(output_folder, f"{row['ID']}.jpg")
            try:
                card.save(image_path, "JPEG")
                images.append(image_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save image for ID {row['ID']}: {e}")

    if images:
        output_pdf_path = os.path.join(output_folder, "output.pdf")
        pdf_path = create_pdf(images, output_pdf_path)
        if pdf_path:
            messagebox.showinfo("Info", f"PDF created successfully: {pdf_path}")
            browse_pdf()
        else:
            messagebox.showerror("Error", "Failed to create PDF.")
    else:
        messagebox.showwarning("Warning", "No valid images to create ID cards.")



root = tk.Tk()
root.title("ID Card Generator")
root.geometry("1200x1200")

try:
    strip_image = Image.open("C:\\Users\\Shree\\Desktop\\riti\\csir_ceeri_logo.png")
    strip_image = strip_image.resize((root.winfo_screenwidth(), 100))  # Resize to fit the width of the screen and set height
    strip_image_tk = ImageTk.PhotoImage(strip_image)

    # Create a frame for the image strip
    strip_frame = tk.Frame(root, bg="white")
    strip_frame.pack(side=tk.TOP, fill=tk.X)

    # Add the image strip to the frame
    strip_label = tk.Label(strip_frame, image=strip_image_tk)
    strip_label.image = strip_image_tk  # Keep a reference to avoid garbage collection
    strip_label.pack(fill=tk.X)
except Exception as e:
    print(f"Error loading image strip: {str(e)}")


    # Create a frame for the image strip
    strip_frame = tk.Frame(root, bg="white")
    strip_frame.pack(side=tk.TOP, fill=tk.X)

    # Add the image strip to the frame
    strip_label = tk.Label(strip_frame, image=strip_image_tk)
    strip_label.image = strip_image_tk  # Keep a reference to avoid garbage collection
    strip_label.pack(fill=tk.X)
except Exception as e:
    print(f"Error loading image strip: {str(e)}")

# Create the top frame for controls
top_frame = tk.Frame(root, bg="white")
top_frame.pack(side=tk.TOP, fill=tk.X)

# Main frame for content
main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True)
# Create frames for PDF and CSV sections
pdf_frame = tk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=2)
pdf_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

csv_frame = tk.Frame(main_frame)
csv_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Define treeview_frame inside csv_frame
treeview_frame = tk.Frame(csv_frame)
treeview_frame.pack(fill=tk.BOTH, expand=True)


# Initialize the menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)
# Initialize the menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# Define and pack menus
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Select CSV", command=select_csv)
file_menu.add_command(label="Browse and Display PDF", command=browse_pdf)
file_menu.add_command(label="Make Editable", command=lambda: make_editable(treeview, treeview.selection()[0] if treeview.selection() else None))

file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)
generate_menu = tk.Menu(menu_bar, tearoff=0)
generate_menu.add_command(label="Generate All ID Cards", command=lambda: create_id_cards(template_path, image_folder, qr_folder, generation_type="all"))
generate_menu.add_command(label="Generate Individual ID Card", command=lambda: create_id_cards(template_path, image_folder, qr_folder, generation_type="individual"))
generate_menu.add_command(label="Generate Selected ID Cards", command=on_generate_id_cards)

menu_bar.add_cascade(label="Generate", menu=generate_menu)

# Create the PDF menu
pdf_menu = tk.Menu(menu_bar, tearoff=0)
pdf_menu.add_command(label="View PDF", command=browse_pdf)

#pdf_menu.add_command(label="Download PDF", command=download_pdf) # Assuming you need to define download_pdf
menu_bar.add_cascade(label="PDF", menu=pdf_menu)
csv_frame = tk.Frame(root)
csv_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
root.mainloop()
