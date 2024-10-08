import os
import csv
import xmlrpc.client
import tkinter as tk
from tkinter import messagebox, filedialog


# Hàm để xuất dữ liệu sản phẩm ra file CSV
def export_products():
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

                    # Ghi dữ liệu sản phẩm
                    for product in products:
                        category_name = product['categ_id'][1] if product['categ_id'] else 'No Category'
                        writer.writerow([product['name'], product['sale_ok'], product['purchase_ok'],
                                         product['list_price'], product['standard_price'], product['default_code'],
                                         product['detailed_type'], product['invoice_policy'], category_name])

                messagebox.showinfo("Thành công", f"Dữ liệu sản phẩm đã được lưu vào {file_path}")
            else:
                messagebox.showwarning("Hủy bỏ", "Lưu file đã bị hủy.")

        else:
            messagebox.showerror("Lỗi", "Không thể xác thực. Vui lòng kiểm tra thông tin đăng nhập.")
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

# Nút để xuất dữ liệu
export_button = tk.Button(root, text="Export Products", command=export_products)
export_button.grid(row=4, columnspan=2, pady=20)

# Chạy vòng lặp chính của GUI
root.mainloop()
