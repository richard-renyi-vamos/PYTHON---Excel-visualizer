import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def visualize_excel(file_path, sheet_name=0):
    # Load the Excel sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print("Excel Data Preview:")
    print(df.head())  # Print first 5 rows
    
    # Select numerical columns
    numerical_cols = df.select_dtypes(include=['number']).columns
    if numerical_cols.empty:
        print("No numerical data found for visualization.")
        return
    
    # Plot each numerical column
    for col in numerical_cols:
        plt.figure(figsize=(10, 5))
        sns.histplot(df[col], kde=True, bins=20)
        plt.title(f'Distribution of {col}')
        plt.xlabel(col)
        plt.ylabel("Frequency")
        plt.show()
    
    # Correlation heatmap
    if len(numerical_cols) > 1:
        plt.figure(figsize=(8, 6))
        sns.heatmap(df[numerical_cols].corr(), annot=True, cmap='coolwarm', fmt='.2f')
        plt.title('Correlation Heatmap')
        plt.show()

if __name__ == "__main__":
    file_path = input("Enter the path to your Excel file: ")
    visualize_excel(file_path)
