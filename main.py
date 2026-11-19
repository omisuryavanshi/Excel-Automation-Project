
import os
import pandas as pd

# Create output folder if it doesn't exist
os.makedirs("output", exist_ok=True)

# Step 1: Read Excel file
# Ask user for Excel filename
file_name = input("Enter Excel file name: ")

file_path = os.path.join("input", file_name)

# Check whether file exists
if not os.path.exists(file_path):
    print("Error: Excel file not found!")
    exit()

# Read Excel file
df = pd.read_excel(file_path)

# Validate required columns
required_columns = ["Customer", "Product", "Quantity", "Price"]

if not all(col in df.columns for col in required_columns):
    print("Error: Excel file has incorrect columns!")
    exit()

print("Excel file loaded successfully!")

print("Original Data:")
print(df)

# Step 2: Remove duplicates
df = df.drop_duplicates()

# Step 3: Handle missing values
df = df.dropna()

# Step 4: Calculate total sales
df["Total Sales"] = df["Quantity"] * df["Price"]

# Step 5: Product-wise sales report
report = df.groupby("Product")["Total Sales"].sum()

# Step 6: Save cleaned data
df.to_excel("output/cleaned_sales.xlsx", index=False)

# Step 7: Save report
report.to_excel("output/product_report.xlsx")

print("Automation completed successfully!")



# Step 8: Generate Sales Summary

total_revenue = df["Total Sales"].sum()

total_quantity = df["Quantity"].sum()

total_customers = df["Customer"].nunique()

best_product = (
    df.groupby("Product")["Quantity"]
    .sum()
    .idxmax()
)

average_order = df["Total Sales"].mean()

summary = pd.DataFrame({
    "Metric": [
        "Total Revenue",
        "Total Products Sold",
        "Total Unique Customers",
        "Best Selling Product",
        "Average Order Value"
    ],
    "Value": [
        total_revenue,
        total_quantity,
        total_customers,
        best_product,
        round(average_order, 2)
    ]
})

summary.to_excel("output/sales_summary.xlsx", index=False)

print("Sales Summary:")
print(summary)



from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import os

# Format all generated Excel files
for filename in os.listdir("output"):

    if filename.endswith(".xlsx"):

        filepath = os.path.join("output", filename)
        workbook = load_workbook(filepath)

        for sheet in workbook.worksheets:

            # Header formatting
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(
                    start_color="2563EB",
                    end_color="2563EB",
                    fill_type="solid"
                )

            # Adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter

                for cell in column:
                    if cell.value is not None:
                        max_length = max(
                            max_length, len(str(cell.value))
                        )

                sheet.column_dimensions[column_letter].width = max_length + 5

            # Add filters and freeze header
            sheet.auto_filter.ref = sheet.dimensions
            sheet.freeze_panes = "A2"

        workbook.save(filepath)

print("Excel formatting completed successfully!")
