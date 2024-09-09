import os
import csv
import xmlrpc.client

# Thông tin kết nối Odoo
url = "http://localhost:8069"
db = "odoo17e_dev"
username = "admin"
password = "admin"

# Xác thực và kết nối với Odoo
common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
uid = common.authenticate(db, username, password, {})
models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

# Lấy dữ liệu sản phẩm
products = models.execute_kw(db, uid, password, 'product.product', 'search_read', [[]],
                             {'fields': ['name', 'list_price', 'categ_id'],
                              'limit': 10})  # Sử dụng 'limit' để giới hạn số lượng sản phẩm

# Đường dẫn thư mục 'output'
output_dir = os.path.join(os.getcwd(), 'output')

# Kiểm tra và tạo thư mục 'output' nếu chưa tồn tại
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Đường dẫn đầy đủ đến file CSV
csv_file = os.path.join(output_dir, "products.csv")

# Tạo file CSV và ghi dữ liệu vào
with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)

    # Ghi tiêu đề cột
    writer.writerow(['Product Name', 'Price', 'Category'])

    # Ghi dữ liệu sản phẩm
    for product in products:
        category_name = product['categ_id'][1] if product['categ_id'] else 'No Category'
        writer.writerow([product['name'], product['list_price'], category_name])

print(f"Dữ liệu sản phẩm đã được lưu vào {csv_file}")
