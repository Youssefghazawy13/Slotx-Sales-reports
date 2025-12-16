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
        
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

def remove_refunds_and_original_sales(sales_df, debug_mode=False):
    """
    Remove refund transactions AND their corresponding original sales
    """
    if 'quantity' not in sales_df.columns or 'barcode' not in sales_df. columns:
        return sales_df
    
    original_count = len(sales_df)
    
    # Find all refunds (negative quantity)
    refunds = sales_df[sales_df['quantity'] < 0]. copy()
    
    if len(refunds) == 0:
        if debug_mode:
            st.write("✅ **No refunds found**")
        return sales_df
    
    if debug_mode:
        st. write(f"🔍 **Found {len(refunds)} refund transactions**")
    
    # Track indices to remove
    indices_to_remove = set()
    
    # Add all refund indices
    indices_to_remove.update(refunds.index.tolist())
    
    # For each refund, find and mark the original sale
    for idx, refund_row in refunds.iterrows():
        barcode = refund_row. get('barcode')
        refund_qty = abs(refund_row. get('quantity', 0))
        brand = refund_row.get('brand')
        
        # Find matching original sales
        potential_originals = sales_df[
            (sales_df['barcode'] == barcode) &
            (sales_df['brand'] == brand) &
            (sales_df['quantity'] > 0) &
            (~sales_df. index.isin(indices_to_remove))
        ]
        
        # Match by quantity
        matching_sales = potential_originals[potential_originals['quantity'] == refund_qty]
        
        if len(matching_sales) > 0:
            # Remove the first matching sale
            indices_to_remove.add(matching_sales.index[0])
    
    # Remove all marked transactions
    cleaned_df = sales_df[~sales_df.index.isin(indices_to_remove)].copy()
    
    removed_count = original_count - len(cleaned_df)
    
    if debug_mode:
        st.write(f"🚫 **Removed {removed_count} transactions** ({len(refunds)} refunds + {removed_count - len(refunds)} original sales)")
    
    return cleaned_df

def get_best_selling_size(sales_data):
    """Extract and find the best selling size from product names"""
    if len(sales_data) == 0 or 'name_ar' not in sales_data.columns:
        return ''
    
    # Extract sizes from product names (last part after last hyphen)
    size_sales = {}
    
    for _, row in sales_data.iterrows():
        product_name = str(row. get('name_ar', ''))
        quantity = row.get('quantity', 0)
        
        # Extract size (last part after last hyphen)
        if '-' in product_name: 
            size = product_name.split('-')[-1]. strip()
            
            # Add to size sales count
            if size: 
                size_sales[size] = size_sales.get(size, 0) + quantity
    
    # Find the best selling size
    if size_sales:
        best_size = max(size_sales, key=size_sales.get)
        return best_size
    
    return ''

def get_best_selling_products(sales_data):
    """Find the best selling product(s)"""
    if len(sales_data) == 0 or 'name_ar' not in sales_data.columns:
        return ''
    
    # Count sales per product
    product_sales = {}
    
    for _, row in sales_data.iterrows():
        product_name = str(row.get('name_ar', ''))
        quantity = row.get('quantity', 0)
        
        if product_name: 
            product_sales[product_name] = product_sales.get(product_name, 0) + quantity
    
    if not product_sales:
        return ''
    
    # Find the maximum sales quantity
    max_sales = max(product_sales. values())
    
    # Get all products with max sales
    best_products = [product for product, sales in product_sales.items() if sales == max_sales]
    
    # Format output
    if len(best_products) == 1:
        return best_products[0]
    else: 
        # Multiple best sellers
        return ', '.join(best_products)

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
        cell.font = Font(bold=True)
    
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

def create_report_sheet(wb, brand_name, sales_data, inventory_data, payout_cycle):
    """Create Report sheet for a specific brand"""
    ws = wb.create_sheet(f"{brand_name} Report")
    
    # Get branch name (use first occurrence)
    branch_name = sales_data. iloc[0].get('branch_name', '') if len(sales_data) > 0 else ''
    
    # Calculate totals
    total_inventory_qty = inventory_data.get('available_quantity', pd.Series([0])).sum()
    total_inventory_value = (inventory_data.get('available_quantity', pd.Series([0])) * 
                            inventory_data.get('sale_price', pd.Series([0]))).sum()
    total_sales_qty = sales_data.get('quantity', pd.Series([0])).sum()
    total_sales_money = sales_data.get('total', pd.Series([0])).sum()
    
    # Get best selling size and products
    best_size = get_best_selling_size(sales_data)
    best_products = get_best_selling_products(sales_data)
    
    # Build report data
    report_data = [
        ['Branch Name:', branch_name],
        ['', ''],
        ['Brand Name:', brand_name],
        ['', ''],
        ['Brand Deal:', ''],
        ['', ''],
        ['Payout Period:', payout_cycle],
        ['', ''],
        ['Best Selling Size:', best_size],
        ['Best Selling Product:', best_products],
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
    
    # Auto-fit columns
    auto_fit_columns(ws)

def process_files(sales_df, inventory_df, payout_cycle, debug_mode=False):
    """Process the sales and inventory files and generate brand reports"""
    
    # Step 1: Clean column names ONLY
    sales_df. columns = sales_df.columns.str.strip()
    inventory_df.columns = inventory_df.columns.str.strip()
    
    if debug_mode:
        st. write(f"📊 **Total Sales Rows (raw):** {len(sales_df)}")
        st.write(f"📊 **Sales Columns:** {list(sales_df.columns)}")
    
    # Step 2: Remove completely empty rows
    sales_df = sales_df.dropna(how='all')
    inventory_df = inventory_df.dropna(how='all')
    
    # Step 3: Remove refunds AND their original sales
    sales_df = remove_refunds_and_original_sales(sales_df, debug_mode=debug_mode)
    
    if debug_mode:
        st. write(f"📊 **Total Sales Rows (after cleaning):** {len(sales_df)}")
    
    # Step 4: Get unique brands from sales data
    brands = sales_df['brand'].dropna().unique()
    
    if debug_mode:
        st.write(f"📊 **Unique Brands Found:** {len(brands)}")
        for brand in brands:
            brand_count = len(sales_df[sales_df['brand'] == brand])
            st.write(f"  - **{brand}:** {brand_count} sales")
    
    # Create a BytesIO object to store the ZIP file
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for brand in brands:
            # Filter data for this brand
            brand_sales = sales_df[sales_df['brand'] == brand]. copy()
            brand_inventory = inventory_df[inventory_df['brand'] == brand]. copy()
            
            if debug_mode:
                st.write(f"🔍 **Processing {brand}:** {len(brand_sales)} sales rows")
            
            # Skip if no sales data
            if len(brand_sales) == 0:
                continue
            
            # Create workbook
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
            
            # Create sheets (pass payout_cycle to report sheet)
            create_sales_details_sheet(wb, brand, brand_sales)
            create_inventory_sheet(wb, brand, brand_inventory)
            create_report_sheet(wb, brand, brand_sales, brand_inventory, payout_cycle)
            
            # Save to BytesIO
            excel_buffer = BytesIO()
            wb.save(excel_buffer)
            excel_data = excel_buffer.getvalue()
            
            # Sanitize brand name
            safe_brand_name = str(brand).replace('/', '-').replace('\\', '-').replace(':', '-').strip()
            
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

# Debug mode toggle
debug_mode = st.checkbox("🔍 Enable Debug Mode (show detailed info)", value=False)

# Payout Cycle Dropdown
st.subheader("📅 Payout Cycle")
payout_cycle = st.selectbox(
    "Select Payout Cycle",
    options=["Payout Cycle 1", "Payout Cycle 2"],
    index=0,
    help="This will appear in the Report sheet under 'Payout Period'"
)

st.divider()

# File uploaders
col1, col2 = st. columns(2)

with col1:
    st.subheader("📈 Sales Sheet")
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
                
                if debug_mode:
                    st.write("### 🔍 Debug Information:")
                
                # Process and create ZIP (pass payout_cycle)
                zip_buffer = process_files(sales_df, inventory_df, payout_cycle, debug_mode=debug_mode)
                
                # Get number of brands
                brands_count = sales_df['brand'].dropna().nunique()
                
                st.success(f"✅ Successfully generated reports for {brands_count} brand(s)!")
                st.info(f"📅 **Payout Cycle:** {payout_cycle}")
                
                # Download button
                st.download_button(
                    label="📥 Download Brands Reports (ZIP)",
                    data=zip_buffer,
                    file_name=f"Brands_Reports_{payout_cycle. replace(' ', '_')}.zip",
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
    1. **Select Payout Cycle** - Choose between Payout Cycle 1 or 2
    2. **Upload Sales Sheet** - Your sales data Excel file
    3. **Upload Inventory Sheet** - Your inventory/stock Excel file
    4. **(Optional)** Enable **Debug Mode** to see detailed processing info
    5. Click **Generate Reports** button
    6. Download the generated ZIP file containing all brand reports
    
    ### What you'll get:
    - Separate Excel file for each brand that has sales
    - Each file contains 3 sheets: 
        - **Sales Details**: All sales data with totals (refunds excluded)
        - **Inventory**: Stock/inventory data
        - **Report**: Summary with calculations, best selling size & products, and selected Payout Cycle
    - All columns are auto-fitted for easy reading
    - Refunds are automatically removed along with their original transactions
    
    ### Best Selling Analysis:
    - **Best Selling Size**: Automatically extracted from product names (last part after hyphen)
    - **Best Selling Product**: The product(s) with highest sales quantity
    - If multiple products have same top sales, all are listed
    
    ### Debug Mode:
    - Shows exactly how many sales each brand has
    - Shows how many refunds and original sales were removed
    - Helps identify if any data is being lost
    
    ### Required Columns:
    **Sales Sheet:**
    - brand, branch_name, name_ar, barcode, quantity, total
    
    **Inventory Sheet:**
    - brand, branch_name, name_en, barcodes, sale_price, available_quantity
    """)

# Footer
st.divider()
st.markdown(
    "<p style='text-align: center; color: #666;'>Made with ❤️ for Slotx</p>",
    unsafe_allow_html=True
)
