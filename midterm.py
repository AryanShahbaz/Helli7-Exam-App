# import tkinter as tk
# from tkinter import filedialog
# from tkinter import messagebox
# from openpyxl import load_workbook
# from docx import Document
# from docx.shared import Pt
# from docx.oxml.ns import qn
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# import os

# def transfer_data():
#     # Excel Choose
#     excel_file = filedialog.askopenfilename(
#         title="انتخاب فایل اکسل : ",
#         filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
#     )
#     if not excel_file:
#         messagebox.showerror("فایل اکسل","فایل اکسل انتخاب نشد")
#         print("فایل اکسل انتخاب نشد.")
#         return

#     # Word Choose
#     word_file = filedialog.askopenfilename(
#         title=" : انتخاب فایل ورد",
#         filetypes=[("Word Files", "*.docx"), ("All Files", "*.*")]
#     )
#     if not word_file:
#         messagebox.showerror("فایل ورد","فایل ورد انتخاب نشد")
#         print("فایل ورد انتخاب نشد.")
#         return

#     #  Select Output Folder
#     output_folder = filedialog.askdirectory(
#         title="انتخاب پوشه ذخیره فایل‌های ورد"
#     )
#     if not output_folder:
#         messagebox.showerror("پوشه خروجی","پوشه خروجی انتخاب نشد.")
#         print("پوشه ذخیره‌سازی انتخاب نشد.")
#         return

#     try:
#         # Load Excel 
#         wb = load_workbook(excel_file, data_only=True)
#         ws = wb.active

#         # Select Row 4 to 33 for transfer
#         for i, row in enumerate(ws.iter_rows(min_row=4, max_row=33, values_only=True)): 
#             if row is None:
#                 continue

#             # St Info
#             name = row[1]  # B : Name
#             class_name = row[3]  # D : Class
#             average = row[46]  # AU : Average
#             rank = row[47]  # AJ : Rank

#             # Copy word for each row(student)
#             new_doc = Document(word_file) 

#             # First Word Table : Name and Class
#             new_table1 = new_doc.tables[0]  
#             if len(new_table1.rows) > 1:
#                 name_cell = new_table1.rows[1].cells[0]
#                 class_cell = new_table1.rows[1].cells[1]
#                 name_cell.text = str(name) if name else ""
#                 class_cell.text = str(class_name) if class_name else ""

#             # Second Word Table : Scores 
#             new_table2 = new_doc.tables[1]
#             for lesson_index in range(7):  # Lessons Count
#                 base_column = 4 + lesson_index * 6  
#                 for sub_index in range(6):  # 6 Column of each lesson
#                     value = row[base_column + sub_index]
#                     new_table2.rows[lesson_index + 1].cells[sub_index + 1].text = str(value) if value else ""

#             # Third Word Table : Average and Rank
#             new_table3 = new_doc.tables[2]
#             if len(new_table3.rows) > 1:
#                 average_cell = new_table3.rows[1].cells[0]
#                 rank_cell = new_table3.rows[1].cells[1]
#                 average_cell.text = str(average) if average else ""
#                 rank_cell.text = str(rank) if rank else ""

#             # Identify Font 
#             for table in [new_table1, new_table2, new_table3]:
#                 for row in table.rows:
#                     for cell in row.cells:
#                         for paragraph in cell.paragraphs:
#                             for run in paragraph.runs:
#                                 run.font.name = "B Nazanin"
#                                 run._element.rPr.rFonts.set(qn('w:eastAsia'), "B Nazanin")
#                                 run.font.size = Pt(18)
#                             paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

#             # Word Output file name
#             student_name = f"{name}_{class_name}.docx"
#             output_file_path = os.path.join(output_folder, student_name)

#             # Save
#             new_doc.save(output_file_path)
#             print(f"فایل ورد برای {name} با موفقیت ذخیره شد: {output_file_path}")

#     except Exception as e:
#         messagebox.showerror("خطا در پردازش فایل ها",f"خطایی در پردازش فایل ها رخ داد\n {e}")
#         print(f"خطا در پردازش فایل‌ها: {e}")

# # UI
# root = tk.Tk()
# root.title("انتقال داده از اکسل به ورد")
# root.geometry("400x200")

# label = tk.Label(root, text="لطفاً فایل‌های اکسل و ورد را انتخاب کنید.", font=("Arial", 12))
# label.pack(pady=20)

# button = tk.Button(root, text="انتقال داده‌ها", command=transfer_data, font=("Arial", 12), bg="lightblue")
# button.pack(pady=20)

# root.mainloop()

# _____________________________________________________________________________________________________
# _____________________________________________________________________________________________________
# _____________________________________________________________________________________________________
# _____________________________________________________________________________________________________
# _____________________________________________________________________________________________________


import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os


def transfer_data():
    # Excel Choose
    excel_file = filedialog.askopenfilename(
        title="انتخاب فایل اکسل : ",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not excel_file:
        messagebox.showerror("فایل اکسل", "فایل اکسل انتخاب نشد")
        print("فایل اکسل انتخاب نشد.")
        return

    # Word Choose
    word_file = filedialog.askopenfilename(
        title=" : انتخاب فایل ورد",
        filetypes=[("Word Files", "*.docx"), ("All Files", "*.*")]
    )
    if not word_file:
        messagebox.showerror("فایل ورد", "فایل ورد انتخاب نشد")
        print("فایل ورد انتخاب نشد.")
        return

    # Select Output Folder
    output_folder = filedialog.askdirectory(
        title="انتخاب پوشه ذخیره فایل‌های ورد"
    )
    if not output_folder:
        messagebox.showerror("پوشه خروجی", "پوشه خروجی انتخاب نشد.")
        print("پوشه ذخیره‌سازی انتخاب نشد.")
        return

    try:
        # Load Excel 
        wb = load_workbook(excel_file, data_only=True)
        ws = wb.active

        # Count Rows for Progress Bar
        total_rows = 33 - 4 + 1  # Rows 4 to 33
        progress["maximum"] = total_rows

        # Select Row 4 to 33 for transfer
        for i, row in enumerate(ws.iter_rows(min_row=4, max_row=33, values_only=True)): 
            if row is None:
                continue

            # Update Progress Bar
            progress["value"] = i + 1
            progress_label.config(text=f"Processing row {i + 1} of {total_rows}")
            root.update_idletasks()

            # St Info
            name = row[1]  # B : Name
            class_name = row[3]  # D : Class
            average = row[46]  # AU : Average
            rank = row[47]  # AJ : Rank

            # Copy word for each row(student)
            new_doc = Document(word_file)

            # First Word Table : Name and Class
            new_table1 = new_doc.tables[0]
            if len(new_table1.rows) > 1:
                name_cell = new_table1.rows[1].cells[0]
                class_cell = new_table1.rows[1].cells[1]
                name_cell.text = str(name) if name else ""
                class_cell.text = str(class_name) if class_name else ""

            # Second Word Table : Scores 
            new_table2 = new_doc.tables[1]
            for lesson_index in range(7):  # Lessons Count
                base_column = 4 + lesson_index * 6  
                for sub_index in range(6):  # 6 Column of each lesson
                    value = row[base_column + sub_index]
                    new_table2.rows[lesson_index + 1].cells[sub_index + 1].text = str(value) if value else ""

            # Third Word Table : Average and Rank
            new_table3 = new_doc.tables[2]
            if len(new_table3.rows) > 1:
                average_cell = new_table3.rows[1].cells[0]
                rank_cell = new_table3.rows[1].cells[1]
                average_cell.text = str(average) if average else ""
                rank_cell.text = str(rank) if rank else ""

            # Identify Font 
            for table in [new_table1, new_table2, new_table3]:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.name = "B Nazanin"
                                run._element.rPr.rFonts.set(qn('w:eastAsia'), "B Nazanin")
                                run.font.size = Pt(18)
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Word Output file name
            student_name = f"{name}_{class_name}.docx"
            output_file_path = os.path.join(output_folder, student_name)

            # Save
            new_doc.save(output_file_path)
            print(f"فایل ورد برای {name} با موفقیت ذخیره شد: {output_file_path}")

        # Show Success Message
        messagebox.showinfo("پایان", "فرایند با موفقیت انجام شد.")
        progress_label.config(text="Completed!")

    except Exception as e:
        messagebox.showerror("خطا در پردازش فایل ها", f"خطایی در پردازش فایل ها رخ داد\n {e}")
        print(f"خطا در پردازش فایل‌ها: {e}")


# UI
root = tk.Tk()
root.title("انتقال داده از اکسل به ورد")
root.geometry("500x300")
root.config(bg="lightblue")

# Title Label
title_label = tk.Label(root, text="انتقال داده از اکسل به ورد", font=("Arial", 16, "bold"), bg="lightblue")
title_label.pack(pady=10)

# Start Button
button = tk.Button(root, text="شروع انتقال داده‌ها", command=transfer_data, font=("Arial", 12), bg="green", fg="white")
button.pack(pady=20)

# Progress Bar
progress = Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

# Progress Label
progress_label = tk.Label(root, text="", font=("Arial", 10), bg="lightblue")
progress_label.pack()

# Run Application
root.mainloop()
