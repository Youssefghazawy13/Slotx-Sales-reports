from flask import Flask, request, render_template, send_file
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import os
import zipfile
from io import BytesIO
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def create_sales_details_sheet(wb, brand_name, sales_data):
    """Create Sales Details sheet for a specific brand"""
    ws = wb.create_sheet(f"{brand_name} Sales Details")
    
    # Headers
    headers = ['Branch Name', 'Brand Name', 'Product Name', 'Barcode', 'Quantity', 'Price']
    ws.append(headers)
    
    # Make headers bold
    for cell in ws[1]:
        cell.font = Font(bold=True)
    
    # Add data
    total_quantity = 0
    total_price = 0
    
    for _, row in sales_data.iterrows():
        ws.append([
            row. get('branch_name', ''),
            row.get('brand', ''),
            row.get('name_ar', ''),
            row.get('barcode', ''),
            row.get('quantity', 0),
            row.get('total', 0)
        ])
        total_quantity += row.get('quantity', 0)
        total_price += row.get('total', 0)
    
    # Add totals row
    total_row = ['', '', '', '', f'Total={total_quantity}', f'Total={total_price}']
    ws.append(total_row)
    
    # Make total row bold
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

def create_inventory_sheet(wb, brand_name, inventory_data):
    """Create Inventory sheet for a specific brand"""
    ws = wb.create_sheet(f"{brand_name} Inventory")
    
    # Headers
    headers = ['Branch Name', 'Brand', 'Product Name', 'Barcodes', 'Product Price', 'Available Quantity']
    ws.append(headers)
    
    # Make headers bold
    for cell in ws[1]:
        cell. font = Font(bold=True)
    
    # Add data
    for _, row in inventory_data.iterrows():
        ws.append([
            row.get('branch_name', ''),
            row.get('brand', ''),
            row.get('name_en', ''),
            row.get('barcodes', ''),
            row.get('sale_price', 0),
            row.get('available_quantity', 0)
        ])

def create_report_sheet(wb, brand_name, sales_data, inventory_data):
    """Create Report sheet for a specific brand"""
    ws = wb.create_sheet(f"{brand_name} Report")
    
    # Get branch name (use first occurrence)
    branch_name = sales_data.iloc[0].get('branch_name', '') if len(sales_data) > 0 else ''
    
    # Calculate totals
    total_inventory_qty = inventory_data.get('available_quantity', pd.Series([0])).sum()
    total_inventory_value = (inventory_data.get('available_quantity', pd.Series([0])) * 
                            inventory_data.get('sale_price', pd.Series([0]))).sum()
    total_sales_qty = sales_data.get('quantity', pd.Series([0])).sum()
    total_sales_money = sales_data.get('total', pd.Series([0])).sum()
    
    # Build report data
    report_data = [
        ['Branch Name:', branch_name],
        ['', ''],
        ['Brand Name:', brand_name],
        ['', ''],
        ['Brand Deal:', ''],
        ['', ''],
        ['Payout Period', ''],
        ['', ''],
        ['Best Selling Size:', ''],
        ['Best Selling Product:', ''],
        ['', ''],
        ['Total Brand Inventory Quantities:', total_inventory_qty],
        ['Total Brand Inventory Stock Price:', total_inventory_value],
        ['', ''],
        ['Total Sales (Products Quantities):', total_sales_qty],
        ['Total sales (Money):', total_sales_money],
        ['Total Sales After Percentage', ''],
        ['Total Sales After Rent:', '']
    ]
    
    # Add data
    for row_data in report_data:
        ws.append(row_data)
    
    # Make all labels in column A bold
    for row in ws. iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            if cell.value:
                cell.font = Font(bold=True)

def process_files(sales_file, inventory_file):
    """Process the sales and inventory files and generate brand reports"""
    
    # Read Excel files
    sales_df = pd.read_excel(sales_file)
    inventory_df = pd. read_excel(inventory_file)
    
    # Normalize column names (strip whitespace, lowercase)
    sales_df.columns = sales_df.columns.str.strip()
    inventory_df.columns = inventory_df.columns.str. strip()
    
    # Get unique brands from sales data
    brands = sales_df['brand'].dropna().unique()
    
    # Create a BytesIO object to store the ZIP file
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for brand in brands:
            # Filter data for this brand
            brand_sales = sales_df[sales_df['brand'] == brand]
            brand_inventory = inventory_df[inventory_df['brand'] == brand]
            
            # Create workbook
            wb = Workbook()
            # Remove default sheet
            wb.remove(wb. active)
            
            # Create sheets
            create_sales_details_sheet(wb, brand, brand_sales)
            create_inventory_sheet(wb, brand, brand_inventory)
            create_report_sheet(wb, brand, brand_sales, brand_inventory)
            
            # Save to BytesIO
            excel_buffer = BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            # Add to ZIP
            zip_file.writestr(f"{brand}. xlsx", excel_buffer.getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST': 
        # Check if files are present
        if 'sales_file' not in request.files or 'inventory_file' not in request.files:
            return 'Missing files', 400
        
        sales_file = request.files['sales_file']
        inventory_file = request.files['inventory_file']
        
        # Validate files
        if sales_file.filename == '' or inventory_file.filename == '': 
            return 'No files selected', 400
        
        if not (allowed_file(sales_file. filename) and allowed_file(inventory_file.filename)):
            return 'Invalid file type.  Please upload Excel files (. xlsx or .xls)', 400
        
        try:
            # Process files
            zip_buffer = process_files(sales_file, inventory_file)
            
            # Send ZIP file
            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name='Brands_Reports. zip'
            )
        except Exception as e:
            return f'Error processing files: {str(e)}', 500
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
