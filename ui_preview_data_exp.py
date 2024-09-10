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
            products = models.execute_kw(db, uid, password, 'product.product', 'search_read', [[]],
                                         {'fields': ['name', 'sale_ok', 'purchase_ok', 'list_price', 'standard_price',
                                                     'default_code', 'detailed_type', 'invoice_policy', 'categ_id']})

            # Xóa dữ liệu cũ trong bảng nếu có
            for row in table.get_children():
                table.delete(row)

            # Hiển thị dữ liệu sản phẩm trong bảng
            for product in products:
                category_name = product['categ_id'][1] if product['categ_id'] else 'No Category'
                table.insert('', 'end', values=(product['name'], product['sale_ok'], product['purchase_ok'],
                                                product['list_price'], product['standard_price'], product['default_code'],
                                                product['detailed_type'], product['invoice_policy'], category_name))

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
                writer.writerow(['Tên SP', 'Có thể bán', 'Có thể mua', 'Giá bán', 'Chi phí',
                                 'Mã nội bộ', 'Loại sản phẩm', 'Chính sách xuất HĐ', 'Danh mục SP'])

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

# URL Odoo
tk.Label(root, text="Odoo URL:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.E)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=0, column=1, padx=10, pady=5)

# Database
tk.Label(root, text="Database:").grid(row=1, column=0, padx=10, pady=5, sticky=tk.E)
db_entry = tk.Entry(root, width=50)
db_entry.grid(row=1, column=1, padx=10, pady=5)

# Username
tk.Label(root, text="Username:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.E)
username_entry = tk.Entry(root, width=50)
username_entry.grid(row=2, column=1, padx=10, pady=5)

# Password
tk.Label(root, text="Password:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.E)
password_entry = tk.Entry(root, show="*", width=50)
password_entry.grid(row=3, column=1, padx=10, pady=5)

# Nút để tải dữ liệu sản phẩm
load_button = tk.Button(root, text="Load Products", command=load_products)
load_button.grid(row=4, columnspan=2, pady=10)

# Bảng hiển thị dữ liệu sản phẩm
columns = ['Tên SP', 'Có thể bán', 'Có thể mua', 'Giá bán', 'Chi phí', 'Mã nội bộ', 'Loại sản phẩm', 'Chính sách xuất HĐ', 'Danh mục SP']
table = ttk.Treeview(root, columns=columns, show='headings', height=10)

for col in columns:
    table.heading(col, text=col)
    table.column(col, width=120)

table.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

# Nút để xuất dữ liệu ra CSV
export_button = tk.Button(root, text="Export to CSV", command=export_products)
export_button.grid(row=6, columnspan=2, pady=10)

# Chạy vòng lặp chính của GUI
root.mainloop()
