import pandas as pd
from datetime import datetime, timedelta

# Today's date
today = datetime.today().date()

# Sample test data
data = {
    "Name": ["Alice", "Bob", "Charlie", "Diana", "Eve", "Frank", "Grace"],
    "Address": ["Addr1", "Addr2", "Addr3", "Addr4", "Addr5", "Addr6", "Addr7"],
    "Default BF": ["Idli", "Dosa", "Upma", "Poha", "Paratha", "Bread", "Oats"],
    "Today BF": ["", "-", "-2", "Pancakes", "no", "-1", ""],
    "Skip BF Until": ["", "", "", "", "", "", (today + timedelta(days=1)).isoformat()],
    "Default Lunch": ["Rice", "Dal", "Curry", "Roti", "Biryani", "Khichdi", "Pasta"],
    "Today Lunch": ["", "Burger", "", "-", "-3", "", ""],
    "Skip Lunch Until": ["", "", "", "", "", (today + timedelta(days=2)).isoformat(), ""],
    "Default Dinner": ["Soup", "Chole", "Pizza", "Samosa", "Rajma", "Noodles", "Salad"],
    "Today Dinner": ["-1", "", "Tacos", "no", "", "", ""],
    "Skip Dinner Until": ["", "", "", "", (today + timedelta(days=3)).isoformat(), "", ""]
}

df = pd.DataFrame(data)

# Save the file
output_filename = "meal_test_cases.xlsx"
df.to_excel(output_filename, index=False)

print(f"Excel file '{output_filename}' has been created with all test cases.")
