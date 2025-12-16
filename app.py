import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import zipfile
from io import BytesIO

st.set_page_config(
    page_title="Slotx Sales & Inventory Reports",
    page_icon="📊",
    layout="centered"
)

def auto_fit_columns(ws):
    """Auto-fit all columns in the worksheet based on content"""
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        
        for cell in column:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length: 
                        max_length = cell_length
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)  # Max width 50
        ws.column_dimensions[column_letter].width = adjusted_width

def create_sales_details_sheet(wb, brand_name, sales_data):
    """Create Sales Details sheet for a specific brand"""
    ws = wb.create_sheet(f"{brand_name} Sales Details")
    
    # Headers
    headers = ['Branch Name', 'Brand Name', 'Product Name', 'Barcode', 'Quantity', 'Price']
    ws. append(headers)
    
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
    
    # Auto-fit columns
    auto_fit_columns(ws)

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
    
    # Auto-fit columns
    auto_fit_columns(ws)

def create_report_sheet(wb, brand_name, sales_data, inventory_data):
    """Create Report sheet for a specific brand"""
    ws = wb.create_sheet(f"{brand_name} Report")
    
    # Get branch name (use first occurrence)
    branch_name = sales_data.iloc[0].get('branch_name', '') if len(sales_data) > 0 else ''
    
    # Calculate totals
    total_inventory_qty = inventory_data. get('available_quantity', pd.Series([0])).sum()
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
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            if cell.value:
                cell.font = Font(bold=True)
    
    # Auto-fit columns
    auto_fit_columns(ws)

def process_files(sales_df, inventory_df):
    """Process the sales and inventory files and generate brand reports"""
    
    # Normalize column names (strip whitespace)
    sales_df. columns = sales_df.columns. str.strip()
    inventory_df.columns = inventory_df. columns.str.strip()
    
    # Get unique brands from sales data only
    brands = sales_df['brand'].dropna().unique()
    
    # Create a BytesIO object to store the ZIP file
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile. ZIP_DEFLATED) as zip_file:
        for brand in brands:
            # Filter data for this brand
            brand_sales = sales_df[sales_df['brand'] == brand]
            brand_inventory = inventory_df[inventory_df['brand'] == brand]
            
            # Create workbook
            wb = Workbook()
            # Remove default sheet
            if 'Sheet' in wb. sheetnames:
                wb. remove(wb['Sheet'])
            
            # Create sheets
            create_sales_details_sheet(wb, brand, brand_sales)
            create_inventory_sheet(wb, brand, brand_inventory)
            create_report_sheet(wb, brand, brand_sales, brand_inventory)
            
            # Save to BytesIO with proper handling
            excel_buffer = BytesIO()
            wb.save(excel_buffer)
            
            # Get the value BEFORE closing
            excel_data = excel_buffer.getvalue()
            
            # Sanitize brand name for filename
            safe_brand_name = str(brand).replace('/', '-').replace('\\', '-').strip()
            
            # Add to ZIP
            zip_file. writestr(f"{safe_brand_name}.xlsx", excel_data)
            
            # Close buffers
            excel_buffer.close()
            wb.close()
    
    zip_buffer.seek(0)
    return zip_buffer

# Streamlit UI
st.title("📊 Slotx Sales & Inventory Reports Generator")
st.markdown("Generate brand-specific sales and inventory reports from Excel files")

st.divider()

# File uploaders
col1, col2 = st.columns(2)

with col1:
    st. subheader("📈 Sales Sheet")
    sales_file = st.file_uploader(
        "Upload Sales Excel File",
        type=['xlsx', 'xls'],
        key='sales',
        help="Upload the sales data Excel file"
    )

with col2:
    st. subheader("📦 Inventory Sheet")
    inventory_file = st.file_uploader(
        "Upload Inventory Excel File",
        type=['xlsx', 'xls'],
        key='inventory',
        help="Upload the inventory data Excel file"
    )

st.divider()

# Process button
if sales_file and inventory_file:
    if st.button("🚀 Generate Reports", type="primary", use_container_width=True):
        try:
            with st.spinner("Processing files...  Please wait"):
                # Read Excel files
                sales_df = pd.read_excel(sales_file)
                inventory_df = pd. read_excel(inventory_file)
                
                # Process and create ZIP
                zip_buffer = process_files(sales_df, inventory_df)
                
                # Get number of brands
                brands_count = sales_df['brand'].dropna().nunique()
                
                st.success(f"✅ Successfully generated reports for {brands_count} brand(s)!")
                
                # Download button
                st.download_button(
                    label="📥 Download Brands Reports (ZIP)",
                    data=zip_buffer,
                    file_name="Brands_Reports.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            st.exception(e)
else:
    st.info("ℹ️ Please upload both Sales and Inventory Excel files to continue")

# Instructions
with st.expander("📖 Instructions"):
    st.markdown("""
    ### How to use:
    1. **Upload Sales Sheet** - Your sales data Excel file
    2. **Upload Inventory Sheet** - Your inventory/stock Excel file
    3. Click **Generate Reports** button
    4. Download the generated ZIP file containing all brand reports
    
    ### What you'll get:
    - Separate Excel file for each brand that has sales
    - Each file contains 3 sheets: 
        - **Sales Details**: Sales data with totals
        - **Inventory**:  Stock/inventory data
        - **Report**: Summary with calculations
    - All columns are auto-fitted for easy reading
    
    ### Required Columns:
    **Sales Sheet:**
    - Column G: barcode
    - Column H: name_ar (Product Name)
    - Column J: brand
    - Column L: quantity
    - Column O: total (Price)
    - Column V: branch_name
    
    **Inventory Sheet:**
    - Column C: name_en (Product Name)
    - Column E: branch_name
    - Column F:  barcodes
    - Column G: brand
    - Column J: sale_price
    - Column N: available_quantity
    """)

# Footer
st.divider()
st.markdown(
    "<p style='text-align: center; color: #666;'>Made with ❤️ for Slotx</p>",
    unsafe_allow_html=True
)
