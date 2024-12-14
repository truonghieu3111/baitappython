import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv
import pandas as pd


CSV_FILE = "nhan_vien.csv"

def save_to_csv(data):
    with open(CSV_FILE, mode="a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(data)

def get_employees_with_birthday_today():
    today = datetime.now().strftime("%d/%m/%Y")
    employees = []
    try:
        with open(CSV_FILE, mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            next(reader)  
            for row in reader:
                if row[3] == today:  
                    employees.append(row)
    except FileNotFoundError:
        pass
    return employees

def export_to_excel():
    try:
        df = pd.read_csv(CSV_FILE, encoding="utf-8")
        df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y")
        df.sort_values(by="Ngày sinh", ascending=True, inplace=True) 
        df.to_excel("Danh_sach_nhan_vien.xlsx", index=False, encoding="utf-8")
        messagebox.showinfo("Thông báo", "Xuất file Excel thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel: {e}")

def submit_form():
    data = [
        entry_id.get(),
        entry_name.get(),
        combo_unit.get(),
        entry_dob.get(),
        gender_var.get(),
        entry_id_card.get(),
        entry_issue_date.get(),
        entry_issue_place.get()
    ]

    if "" in data:
        messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ thông tin!")
        return

    save_to_csv(data)
    messagebox.showinfo("Thông báo", "Lưu thông tin thành công!")
    clear_form()

def show_birthday_today():
    employees = get_employees_with_birthday_today()
    if not employees:
        messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")
    else:
        result = "\n".join([f"Mã: {e[0]}, Tên: {e[1]}" for e in employees])
        messagebox.showinfo("Danh sách sinh nhật hôm nay", result)

def clear_form():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    combo_unit.set("")
    entry_dob.delete(0, tk.END)
    gender_var.set("Nam")
    entry_id_card.delete(0, tk.END)
    entry_issue_date.delete(0, tk.END)
    entry_issue_place.delete(0, tk.END)


root = tk.Tk()
root.title("Thông tin nhân viên")


label_id = ttk.Label(root, text="Mã *")
label_id.grid(row=0, column=0, padx=5, pady=5)
entry_id = ttk.Entry(root)
entry_id.grid(row=0, column=1, padx=5, pady=5)

label_name = ttk.Label(root, text="Tên *")
label_name.grid(row=0, column=2, padx=5, pady=5)
entry_name = ttk.Entry(root)
entry_name.grid(row=0, column=3, padx=5, pady=5)

label_unit = ttk.Label(root, text="Đơn vị")
label_unit.grid(row=1, column=0, padx=5, pady=5)
combo_unit = ttk.Combobox(root, values=["Phân xưởng quê hương", "Văn phòng", "Khác"])
combo_unit.grid(row=1, column=1, padx=5, pady=5)

label_dob = ttk.Label(root, text="Ngày sinh (DD/MM/YYYY)")
label_dob.grid(row=1, column=2, padx=5, pady=5)
entry_dob = ttk.Entry(root)
entry_dob.grid(row=1, column=3, padx=5, pady=5)

label_gender = ttk.Label(root, text="Giới tính")
label_gender.grid(row=2, column=0, padx=5, pady=5)
gender_var = tk.StringVar(value="Nam")
radio_male = ttk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam")
radio_male.grid(row=2, column=1, padx=5, pady=5)
radio_female = ttk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ")
radio_female.grid(row=2, column=2, padx=5, pady=5)

label_id_card = ttk.Label(root, text="Số CMND")
label_id_card.grid(row=3, column=0, padx=5, pady=5)
entry_id_card = ttk.Entry(root)
entry_id_card.grid(row=3, column=1, padx=5, pady=5)

label_issue_date = ttk.Label(root, text="Ngày cấp")
label_issue_date.grid(row=3, column=2, padx=5, pady=5)
entry_issue_date = ttk.Entry(root)
entry_issue_date.grid(row=3, column=3, padx=5, pady=5)

label_issue_place = ttk.Label(root, text="Nơi cấp")
label_issue_place.grid(row=4, column=0, padx=5, pady=5)
entry_issue_place = ttk.Entry(root)
entry_issue_place.grid(row=4, column=1, padx=5, pady=5)


btn_submit = ttk.Button(root, text="Lưu", command=submit_form)
btn_submit.grid(row=5, column=0, padx=5, pady=5)

btn_birthday = ttk.Button(root, text="Sinh nhật hôm nay", command=show_birthday_today)
btn_birthday.grid(row=5, column=1, padx=5, pady=5)

btn_export = ttk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel)
btn_export.grid(row=5, column=2, padx=5, pady=5)

btn_clear = ttk.Button(root, text="Xóa", command=clear_form)
btn_clear.grid(row=5, column=3, padx=5, pady=5)

root.mainloop()
