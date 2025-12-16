# 📊 Slotx Sales & Inventory Reports Generator

A Flask web application that processes sales and inventory Excel files and generates brand-specific reports.

## Features

- Upload Sales and Inventory Excel files
- Automatically generate separate Excel files for each brand
- Each brand file contains 3 sheets: 
  - **Sales Details**: Sales data for the brand
  - **Inventory**: Stock/inventory data for the brand
  - **Report**: Summary report with calculations
- Download all reports in a single ZIP file

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. Clone the repository:
```bash
git clone https://github.com/Youssefghazawy13/Slotx-Sales-reports.git
cd Slotx-Sales-reports
```

2. Create a virtual environment (optional but recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Start the Flask application:
```bash
python app.py
```

2. Open your browser and navigate to:
```
http://localhost:5000
```

3. Upload your files:
   - Select your **Sales Excel file**
   - Select your **Inventory Excel file**
   - Click **"Generate Reports"**

4. Download the generated `Brands_Reports.zip` file

## File Structure

```
Slotx-Sales-reports/
├── app.py                  # Main Flask application
├── requirements.txt        # Python dependencies
├── templates/
│   └── index.html         # Frontend HTML page
└── README.md              # This file
```

## Input File Requirements

### Sales Sheet Columns
- Column G: `barcode`
- Column H: `name_ar` (Product Name in Arabic)
- Column J: `brand`
- Column L: `quantity`
- Column O: `total` (Price)
- Column V: `branch_name`

### Inventory Sheet Columns
- Column C: `name_en` (Product Name in English)
- Column E: `branch_name`
- Column F: `barcodes`
- Column G: `brand`
- Column J: `sale_price`
- Column N: `available_quantity`

## Output

The application generates a ZIP file containing individual Excel files for each brand. Each Excel file includes:

1. **Sales Details Sheet**: Brand-specific sales data with totals
2. **Inventory Sheet**: Brand-specific inventory data
3. **Report Sheet**: Summary report with:
   - Branch and brand information
   - Inventory totals and values
   - Sales totals
   - Fields for manual entry (Brand Deal, Payout Period, etc.)

## Technologies Used

- **Flask**: Web framework
- **Pandas**: Data processing
- **OpenPyXL**: Excel file manipulation
- **HTML/CSS/JavaScript**: Frontend

## License

MIT License

## Author

Youssefghazawy13
