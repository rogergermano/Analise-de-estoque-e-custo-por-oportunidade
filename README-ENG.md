# Inventory & Opportunity Cost Analyzer (v2.0)
A data science tool focused on inventory financial health. Unlike standard stock reports, this script employs Opportunity Cost and Anti-False Rotation Logic to identify working capital being eroded by time, even for items with recent isolated sales.

# Execution Environment
Important Note: This script is specifically designed to run on Google Colab. It relies on cloud-native libraries for Google service authentication (google.colab, gspread) and assumes a cloud-based file system structure.

# Key Features

- Investment-Based Logic: Evaluates how long capital has been "trapped" in an item, preventing single sales from masking products stagnant for months or years.
- Financial Correction: Applies compound interest to the acquisition cost to reveal the updated real profit margin.
- Localized Dashboards: Bar and scatter plots with Brazilian currency formatting (R$) and visual discrepancy handling (clipping extreme margins for better scaling).

# Required Data Structure
The input file (CSV or Excel) must include:

- SKU / Code: Product identifier.
- Stock: Current physical quantity.
- Cost: Original purchase price.
- Price: Current selling price.
- Last Purchase: Batch entry date.
- Last Sale: Last sold date.

# Customization
If your database uses different column names, simply update the mapping within the processar_dados function. The script is sector-agnostic, performing effectively for construction materials, auto parts, retail, and wholesale distributors.
