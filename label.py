import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors



# from reportlab.lib.pagesizes import A4

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

class ExitTypeGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Exit Type Generator")
        self.root.geometry("600x750")
        self.root.configure(bg="#f4f6f9")

        # Register Verdana Bold font
        verdana_path = "C:\\Windows\\Fonts\\verdana.ttf"
        verdana_bold_path = "C:\\Windows\\Fonts\\verdanab.ttf"
        pdfmetrics.registerFont(TTFont("Verdana", verdana_path))
        pdfmetrics.registerFont(TTFont("Verdana-Bold", verdana_bold_path))

        # title = tk.Label(root, text="Exit Type Label Generator", bg="#007acc", fg="white",
        #                  font=("Verdana", 18, "bold"), pady=10)
        # title.pack(fill=tk.X)

        form_frame = tk.Frame(root, bg="#f2f2f2", padx=20, pady=20)
        form_frame.pack(pady=10)

        # Set icon (optional)
        try:
            self.root.iconbitmap("labelicon.ico")
        except tk.TclError:
            print("Icon not found, using default")

        # Title Label
        title_label = tk.Label(
            root, text="Exit Type Generator",
            font=("Arial", 18, "bold"), bg="#2e86de", fg="white", pady=10
        )
        title_label.pack(fill=tk.X)

        # Frame for form inputs
        self.form_frame = tk.Frame(root, bg="#f4f6f9")
        self.form_frame.pack(pady=15)

        # Labels list (excluding Description for special handling)
        labels = [
            "Inspection No", "Exporter", "Inspection Date", "Bill No", "BOE Date",
            "Manifest", "Country of Origin", "Point of Exit", "Destination",
            "Quantity", "Total Quantity", "Total Weight", "Container/Vehicle No"
        ]

        self.entries = {}

        # Create standard entries
        for i, label_text in enumerate(labels):
            tk.Label(self.form_frame, text=label_text + ":", font=("Arial", 10), bg="#f4f6f9") \
                .grid(row=i, column=0, sticky="e", padx=10, pady=5)
            entry = tk.Entry(self.form_frame, width=45)
            entry.grid(row=i, column=1, padx=10, pady=5)
            self.entries[label_text] = entry

        # Special: Multi-line Text Area for Description
        desc_label = tk.Label(self.form_frame, text="Description:", font=("Arial", 10), bg="#f4f6f9")
        desc_label.grid(row=len(labels), column=0, sticky="ne", padx=10, pady=5)
        self.desc_text = tk.Text(self.form_frame, width=45, height=5, font=("Arial", 10))
        self.desc_text.grid(row=len(labels), column=1, padx=10, pady=5)

        # Buttons
        btn_frame = tk.Frame(root, bg="#f4f6f9")
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Generate PDF", command=self.generate_pdf_label,
                  bg="#27ae60", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="Clear All", command=self.clear_fields,
                  bg="#c0392b", fg="white", font=("Arial", 11, "bold"), padx=15, pady=5).grid(row=0, column=1, padx=10)

        # Text output area
        tk.Label(root, text="Exit Label Preview:", font=("Arial", 11, "bold"), bg="#f4f6f9").pack(pady=(10, 0))
        self.preview_text = tk.Text(root, height=10, width=70, font=("Courier New", 10))
        self.preview_text.pack(padx=10, pady=5)

    def clear_fields(self):
        """Clear all input fields and preview"""
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.desc_text.delete('1.0', tk.END)
        self.preview_text.delete('1.0', tk.END)

    def generate_pdf_label(self):
        """Collect data, display preview, and create PDF"""
        data = {label: entry.get().strip() for label, entry in self.entries.items()}
        data["Description"] = self.desc_text.get("1.0", tk.END).strip()

        # if not all(data.values()):
        #     messagebox.showwarning("Missing Info", "Please fill all fields before generating.")
        #     return

        # Preview text
        preview_lines = [f"{label}: {value}" for label, value in data.items()]
        preview_text = "\n".join(preview_lines)
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert(tk.END, preview_text)

        # Ask where to save
        pdf_filename = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
            title="Save Exit Type PDF As"
        )
        if not pdf_filename:
            return

        # Generate the PDF
        self.create_pdf_label(data, pdf_filename)
        messagebox.showinfo("Success", f"Exit Type PDF saved as:\n{pdf_filename}")

    def create_pdf_label(self, data, filename):
        """Generate the label PDF with precise field positioning (x, y control)"""
        c = canvas.Canvas(filename, pagesize=letter)
        width, height = letter

        # Register Verdana Bold font for title
        c.setFont("Verdana-Bold", 16)
        c.setFillColor(colors.darkblue)
        c.drawString(1 * inch, height - 1.2 * inch, "Exit Type Label")
        c.setFillColor(colors.black)
        c.setFont("Verdana", 11)

        # 🔹 Define X, Y positions for each field
        field_positions = {
            "Inspection No": (1 * inch, height - 1.8 * inch),
            "Exporter": (1 * inch, height - 2.1 * inch),
            "Inspection Date": (1 * inch, height - 2.4 * inch),
            "Bill No": (1 * inch, height - 2.7 * inch),
            "BOE Date": (1 * inch, height - 3.0 * inch),
            "Manifest": (1 * inch, height - 3.3 * inch),
            "Country of Origin": (1 * inch, height - 3.6 * inch),
            "Point of Exit": (4.0 * inch, height - 1.8 * inch),
            "Destination": (4.0 * inch, height - 2.1 * inch),
            "Quantity": (4.0 * inch, height - 2.4 * inch),
            "Total Quantity": (4.0 * inch, height - 2.7 * inch),
            "Total Weight": (4.0 * inch, height - 3.0 * inch),
            "Container/Vehicle No": (4.8 * inch, height - 3.3 * inch),
        }

        # 🔹 Draw fields using the above coordinates
        for key, (x, y) in field_positions.items():
            c.setFont("Verdana-Bold", 11)
            c.drawString(x, y, f"{key}:")
            c.setFont("Verdana", 11)
            c.drawString(x + 1.8 * inch, y, str(data.get(key, "")))

        # 🔹 Description block with multi-line handling
        desc_x = 1 * inch
        desc_y = height - 4.3 * inch
        c.setFont("Verdana-Bold", 11)
        c.drawString(desc_x, desc_y, "Description:")
        c.setFont("Verdana", 10)

        desc_text = data.get("Description", "")
        desc_lines = desc_text.splitlines()

        y_offset = desc_y - 0.3 * inch
        for line in desc_lines:
            c.drawString(desc_x + 0.3 * inch, y_offset, line)
            y_offset -= 0.25 * inch

        # 🔹 Footer
        c.setStrokeColor(colors.grey)
        c.line(1 * inch, 0.9 * inch, width - 1 * inch, 0.9 * inch)
        c.setFont("Verdana", 9)
        c.setFillColor(colors.grey)
        c.drawString(1 * inch, 0.7 * inch, "Generated by Exit Type Generator © 2025")

        c.save()



if __name__ == "__main__":
    root = tk.Tk()
    app = ExitTypeGeneratorApp(root)
    root.mainloop()








# import tkinter as tk
# from tkinter import filedialog, messagebox
# from reportlab.lib.pagesizes import A4
# from reportlab.pdfgen import canvas
# from reportlab.lib.units import inch
# from reportlab.pdfbase import pdfmetrics
# from reportlab.pdfbase.ttfonts import TTFont
# import os

# class ExitTypeGenerator:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("Exit Type Generator")
#         self.root.geometry("700x650")
#         self.root.configure(bg="#f2f2f2")

#         # Register Verdana Bold font
#         verdana_path = "C:\\Windows\\Fonts\\verdana.ttf"
#         verdana_bold_path = "C:\\Windows\\Fonts\\verdanab.ttf"
#         pdfmetrics.registerFont(TTFont("Verdana", verdana_path))
#         pdfmetrics.registerFont(TTFont("Verdana-Bold", verdana_bold_path))

#         title = tk.Label(root, text="Exit Type Label Generator", bg="#007acc", fg="white",
#                          font=("Verdana", 18, "bold"), pady=10)
#         title.pack(fill=tk.X)

#         form_frame = tk.Frame(root, bg="#f2f2f2", padx=20, pady=20)
#         form_frame.pack(pady=10)

#         labels = [
#             "Inspection No", "Exporter", "Inspection Date", "Bill No",
#             "BOE Date", "Manifest", "Country of Origin", "Point of Exit",
#             "Destination", "Quantity", "Description", "Total Quantity",
#             "Total Weight", "Container/Vehicle No"
#         ]

#         self.entries = {}

#         for i, label_text in enumerate(labels):
#             label = tk.Label(form_frame, text=label_text, bg="#f2f2f2", anchor="w",
#                              font=("Verdana", 10, "bold"))
#             label.grid(row=i, column=0, sticky="w", pady=5)

#             if label_text == "Description":
#                 text_area = tk.Text(form_frame, width=50, height=4, font=("Verdana", 10))
#                 text_area.grid(row=i, column=1, pady=5)
#                 self.entries[label_text] = text_area
#             else:
#                 entry = tk.Entry(form_frame, width=53, font=("Verdana", 10))
#                 entry.grid(row=i, column=1, pady=5)
#                 self.entries[label_text] = entry

#         tk.Button(root, text="Generate PDF", command=self.generate_pdf_label,
#                   bg="#007acc", fg="white", font=("Verdana", 12, "bold"),
#                   relief="raised", padx=10, pady=5).pack(pady=20)

#     def generate_pdf_label(self):
#         try:
#             data = {key: (val.get("1.0", tk.END).strip() if isinstance(val, tk.Text) else val.get())
#                     for key, val in self.entries.items()}

#             pdf_filename = filedialog.asksaveasfilename(defaultextension=".pdf",
#                                                         filetypes=[("PDF files", "*.pdf")])
#             if not pdf_filename:
#                 return

#             self.create_pdf_label(**data, pdf_filename=pdf_filename)
#             messagebox.showinfo("Success", f"PDF saved successfully:\n{pdf_filename}")
#         except Exception as e:
#             messagebox.showerror("Error", str(e))

#     def create_pdf_label(self, Inspection_No, Exporter, Inspection_Date, Bill_No,
#                          BOE_Date, Manifest, Country_of_Origin, Point_of_Exit,
#                          Destination, Quantity, Description, Total_Quantity,
#                          Total_Weight, Container_Vehicle_No, pdf_filename):

#         c = canvas.Canvas(pdf_filename, pagesize=A4)
#         c.setFont("Verdana-Bold", 14)
#         c.drawString(1 * inch, 10.5 * inch, "EXIT TYPE LABEL")

#         c.setFont("Verdana", 11)
#         x = 1 * inch
#         y = 10 * inch
#         spacing = 0.35 * inch

#         fields = [
#             ("Inspection No", Inspection_No),
#             ("Exporter", Exporter),
#             ("Inspection Date", Inspection_Date),
#             ("Bill No", Bill_No),
#             ("BOE Date", BOE_Date),
#             ("Manifest", Manifest),
#             ("Country of Origin", Country_of_Origin),
#             ("Point of Exit", Point_of_Exit),
#             ("Destination", Destination),
#             ("Quantity", Quantity),
#             ("Total Quantity", Total_Quantity),
#             ("Total Weight", Total_Weight),
#             ("Container/Vehicle No", Container_Vehicle_No)
#         ]

#         for label, value in fields:
#             c.drawString(x, y, f"{label}: {value}")
#             y -= spacing

#         # Handle multi-line description properly
#         c.setFont("Verdana-Bold", 12)
#         c.drawString(x, y - 0.2 * inch, "Description:")
#         y -= 0.6 * inch
#         c.setFont("Verdana", 11)
#         text_obj = c.beginText(x, y)
#         for line in Description.splitlines():
#             text_obj.textLine(line)
#         c.drawText(text_obj)

#         c.save()

# # Run the app
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = ExitTypeGenerator(root)
#     root.mainloop()
