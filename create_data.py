import os
import pandas as pd

os.makedirs("input", exist_ok=True)

data = {
    "Customer": ["Amit", "Rahul", "Priya", "Amit", "Neha", "Raj"],
    "Product": ["Laptop", "Mouse", "Keyboard", "Laptop", "Mouse", "Monitor"],
    "Quantity": [1, 2, 1, 1, 3, 1],
    "Price": [50000, 500, 1000, 50000, 500, 12000]
}

df = pd.DataFrame(data)

df.to_excel("input/sales.xlsx", index=False)

print("Excel file created successfully!")