import os
import csv
import xmlrpc.client
import tkinter as tk
from tkinter import messagebox, filedialog, ttk


# Hàm để tải dữ liệu sản phẩm và hiển thị trong bảng
def load_products():
    try:
        # Thông tin kết nối Odoo
        url = url_entry.get()
        db = db_entry.get()
        username = username_entry.get()
        password = password_entry.get()

        # Xác thực và kết nối với Odoo
        common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
        uid = common.authenticate(db, username, password, {})
        if uid:
            models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

            # Lấy dữ liệu sản phẩm
            partner = models.execute_kw(db, uid, password, 'res.partner', 'search_read', [[]],
                                         {'fields': ['name', 'is_company', 'company_name', 'country_id', 'state_id',
                                                     'zip', 'street', 'street2', 'phone', 'mobile', 'email', 'vat']})

            # Xóa dữ liệu cũ trong bảng nếu có
            for row in table.get_children():
                table.delete(row)

            # Hiển thị dữ liệu sản phẩm trong bảng
            for partner in partner:
                # category_name = product['categ_id'][1] if product['categ_id'] else 'No Category'
                table.insert('', 'end', values=(partner['name'], partner['is_company'], partner['company_name'],
                                                partner['country_id'], partner['state_id'], partner['zip'],
                                                partner['street'], partner['street2'], partner['phone'],
                                                partner['mobile'], partner['email'], partner['vat']))

            messagebox.showinfo("Thành công", "Dữ liệu sản phẩm đã được tải lên thành công.")
        else:
            messagebox.showerror("Lỗi", "Không thể xác thực. Vui lòng kiểm tra thông tin đăng nhập.")
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


# Hàm để xuất dữ liệu sản phẩm ra file CSV
def export_products():
    try:
        # Hộp thoại để chọn vị trí lưu file
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")],
                                                 title="Chọn vị trí lưu file")

        if file_path:
            # Tạo file CSV và ghi dữ liệu vào
            with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)

                # Ghi tiêu đề cột
                writer.writerow(['Tên', 'Là công ty', 'Tên công ty', 'Quốc gia', 'Tỉnh thành',
                                 'Mã bưu chính', 'Đường', 'Đường 2', 'Điện thoại', 'Di động', 'email', 'vat'])

                # Ghi dữ liệu sản phẩm từ bảng
                for row_id in table.get_children():
                    row = table.item(row_id)['values']
                    writer.writerow(row)

            messagebox.showinfo("Thành công", f"Dữ liệu sản phẩm đã được lưu vào {file_path}")
        else:
            messagebox.showwarning("Hủy bỏ", "Lưu file đã bị hủy.")
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

# Tạo cửa sổ chính
root = tk.Tk()
root.title("Odoo Product Exporter")
root.geometry("800x400")

# URL Odoo
tk.Label(root, text="Odoo URL:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)

# Database
tk.Label(root, text="Database:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
db_entry = tk.Entry(root, width=50)
db_entry.grid(row=1, column=1, padx=10, pady=5, sticky=tk.W)

# Username
tk.Label(root, text="Username:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
username_entry = tk.Entry(root, width=50)
username_entry.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)

# Password
tk.Label(root, text="Password:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
password_entry = tk.Entry(root, show="*", width=50)
password_entry.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)

# Tạo khung để chứa bảng và thanh cuộn
frame = tk.Frame(root)
frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

# Tạo thanh cuộn dọc
scrollbar_y = ttk.Scrollbar(frame, orient='vertical')
scrollbar_y.pack(side='right', fill='y')

# Tạo thanh cuộn ngang
scrollbar_x = ttk.Scrollbar(frame, orient='horizontal')
scrollbar_x.pack(side='bottom', fill='x')

# Tạo bảng hiển thị dữ liệu sản phẩm
columns = ['Tên', 'Là công ty', 'Tên công ty', 'Quốc gia', 'Tỉnh thành',
           'Mã bưu chính', 'Đường', 'Đường 2', 'Điện thoại', 'Di động', 'email', 'vat']
table = ttk.Treeview(frame, columns=columns, show='headings', height=10,
                     yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

# Thiết lập cột và tiêu đề cột
for col in columns:
    table.heading(col, text=col)
    table.column(col, width=120)

table.pack(side='left', fill='both', expand=True)

# Liên kết thanh cuộn với bảng
scrollbar_y.config(command=table.yview)
scrollbar_x.config(command=table.xview)

# Tạo một khung mới để chứa các nút bấm và định vị nó bên phải
button_frame = tk.Frame(root)
button_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='e')

# Nút để tải dữ liệu
load_button = tk.Button(button_frame, text="Load Data")
load_button.pack(side='right', padx=5)

# Nút để xuất dữ liệu ra CSV
export_button = tk.Button(button_frame, text="Export to CSV")
export_button.pack(side='right')

# Đặt tỷ lệ cho cột và hàng để thanh cuộn hoạt động tốt hơn
root.grid_rowconfigure(4, weight=1)
root.grid_columnconfigure(1, weight=1)

# Chạy vòng lặp chính của GUI
root.mainloop()

