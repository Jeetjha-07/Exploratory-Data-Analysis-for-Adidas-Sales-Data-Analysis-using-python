import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

sns.set(style="whitegrid")
plt.rcParams["figure.figsize"] = (12, 6)

# Load CSV dataset
file_path = r"C:\Users\HP\Downloads\PythonProject\Adidas US Sales Datasets.xlsx"
df = pd.read_excel(file_path, sheet_name="Data Sales Adidas", engine='openpyxl')

# Convert dates and numerics
df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
numeric_cols = ['Total Sales', 'Price per Unit', 'Units Sold']
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

df.dropna(subset=numeric_cols, inplace=True)
df.reset_index(drop=True, inplace=True)

print(f"\nData Loaded: {df.shape[0]} rows and {df.shape[1]} columns")
print("\nSummary Statistics:")
print(df.describe())

# ======================= PLOT FUNCTIONS =========================

def plot_bar_by_column(column):
    if column not in df.columns:
        print(f"[ERROR] Column '{column}' not found.")
        return
    top_values = df[column].value_counts().head(10)
    if top_values.empty:
        print(f"[WARNING] Column '{column}' has no data to plot.")
        return
    top_values.plot(kind="bar", color="orange")
    plt.title(f"Top 10 {column}")
    plt.ylabel("Frequency")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def plot_sales_by_group(group_col):
    if group_col not in df.columns:
        print(f"[ERROR] Column '{group_col}' not found.")
        return
    grouped = df.groupby(group_col)['Total Sales'].sum().sort_values(ascending=False)
    grouped = grouped[grouped > 0]
    if grouped.empty:
        print(f"[WARNING] No valid sales data for '{group_col}'.")
        return
    grouped.head(10).plot(kind='bar', color=np.random.choice(['purple', 'skyblue', 'green']))
    plt.title(f'Top {group_col} by Total Sales')
    plt.ylabel('Sales')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def plot_correlation():
    numeric_df = df.select_dtypes(include=np.number).dropna(axis=1, how='all')
    if numeric_df.shape[1] < 2:
        print("[WARNING] Not enough numeric data for correlation heatmap.")
        return
    sns.heatmap(numeric_df.corr(), annot=True, cmap="coolwarm")
    plt.title("Correlation Heatmap")
    plt.tight_layout()
    plt.show()

def plot_pairplot():
    try:
        sns.pairplot(df[['Units Sold', 'Price per Unit', 'Total Sales']], diag_kind='kde')
        plt.suptitle("Pairplot of Numeric Columns", y=1.02)
        plt.tight_layout()
        plt.show()
    except Exception as e:
        print("[ERROR] Could not generate pairplot:", e)

def plot_violin():
    if 'Region' not in df.columns:
        print("[ERROR] Column 'Region' not found for violin plot.")
        return
    sns.violinplot(x='Region', y='Total Sales', data=df, palette='Set2')
    plt.title("Violin Plot: Total Sales by Region")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# ======================= INTERACTIVE MENU =========================

def run_interactive():
    while True:
        print("\n" + "="*60)
        print("ADIDAS SALES DATA ANALYSIS MENU")
        print("="*60)
        print("1. Summary Statistics")
        print("2. Correlation Heatmap")
        print("3. Sales by Region")
        print("4. Sales by State")
        print("5. Sales by Product")
        print("6. Violin Plot: Sales by Region")
        print("7. Top 10 by Gender / Product / Retailer (Choose)")
        print("8. Run All")
        print("0. Exit")
        print("="*60)
        
        choice = input("Enter your choice (0-8): ")

        if choice == '1':
            print(df.describe())
        elif choice == '2':
            plot_correlation()
        elif choice == '3':
            plot_sales_by_group('Region')
        elif choice == '4':
            plot_sales_by_group('State')
        elif choice == '5':
            plot_sales_by_group('Product')
        elif choice == '6':
            plot_violin()
        elif choice == '7':
            col = input("Enter column name (Gender/Product/Retailer): ")
            plot_bar_by_column(col)
        elif choice == '8':
            plot_correlation()
            plot_sales_by_group('Region')
            plot_sales_by_group('State')
            plot_sales_by_group('Product')
            plot_violin()
            plot_pairplot()
            plot_bar_by_column("Retailer")
            plot_bar_by_column("Product")
        elif choice == '0':
            print("\nThank you for using the Adidas Sales EDA Tool!")
            break
        else:
            print("Invalid choice. Try again.")

# ======================= RUN =========================

ans = input("\nDo you want to run interactively? (y/n): ")
if ans.lower() == 'y':
    run_interactive()
else:
    print("\nRunning All Analyses Automatically...")
    plot_correlation()
    plot_sales_by_group('Region')
    plot_sales_by_group('State')
    plot_sales_by_group('Product')
    plot_violin()
    plot_pairplot()
    #plot_bar_by_column("Gender")
    plot_bar_by_column("Retailer")
    plot_bar_by_column("Product")
    print("\nAll EDA Complete!")
