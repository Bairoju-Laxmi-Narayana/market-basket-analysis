# Import necessary libraries
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import openpyxl
product = input(' enter the product          :        ')
# Load the dataset from a CSV file
file_path = r"C:\Users\bombo\OneDrive\Documents\archive\groceries.xlsx"
wb = openpyxl.load_workbook(file_path)

# ws =  wb.sheet_by_index(0)
# for i in range(ws.nrows):
#     for i in range (ws.ncols):
#         print(ws.cell_value(i,j),end="\t")
#     print('')

sheet = wb.active  # You can also use wb['SheetName'] if you know the sheet name

# converting into a dataframe
data = sheet.values
dataset = pd.DataFrame(data)


# Create a list to store the transaction data
transactions = []

# Loop through each row in the dataset
for i in range(0, 7501):
    # Extract items from the current row and convert to string
    # Append the items to the transactions list
    transactions.append([str(dataset.values[i, j]) for j in range(0, 20)])

# Import the apriori function from the apyori library
from apyori import apriori

# Apply the eclat algorithm to find association rules
rules = apriori(
    transactions=transactions,
    min_support=0.001,
    min_confidence=0.09,
    min_lift=2,
    min_length=2,
    max_length=2
)

# Convert the rules generator object into a list
results = list(rules)

# Define a function to extract relevant information from results
# def inspect(results):
#     lhs = [tuple(result[2][0][0])[0] for result in results]  # Left-hand side item
#     rhs = [tuple(result[2][0][1])[0] for result in results]  # Right-hand side item
#     supports = [result[1] for result in results]  # Support value
#     return list(zip(lhs, rhs, supports))

# Assuming the association rules data is stored in a variable called `association_rules`
import pandas as pd

def inspect(results):
    data = []
    for result in results:
        lhs = tuple(result[2][0][0])[0]
        rhs = tuple(result[2][0][1])[0]
        support = result[1]
        lift = result[2][0][2]
        confidence = result[2][0][3]
        data.append([lhs, rhs, support, lift, confidence])
    return data

# Call the inspect function to extract rule details
rule_details = inspect(results)

# Convert the extracted rule details into a DataFrame
resultsinDataFrame = pd.DataFrame(
    rule_details,
    columns=['Product-1', 'Product-2', 'Support','Lift','Confidence']
)
print("resultsinDataFrame")
print(resultsinDataFrame)

# Display the DataFrame containing rule details

unique_items = set()
for itemset in resultsinDataFrame['Product-1']:
    unique_items.add(itemset)
for itemset in resultsinDataFrame['Product-2']:
    unique_items.add(itemset)

print("Unique items:", unique_items)
# Display the top 10 rules with the highest Support values
top_10_support = resultsinDataFrame.nlargest(n=200, columns='Support')
datafr = pd.DataFrame(top_10_support)

# Specify the file path where you want to save the Excel file
file_path = r'C:\Users\bombo\OneDrive\Desktop\eclat.xlsx'

# Use pandas to write the DataFrame to an Excel file
datafr.to_excel(file_path, index=False)


def search_support(resultsinDataFrame, item_1, item_2):
    # Search for the record where item_1 is in 'Product-1' and item_2 is in 'Product-2'
    condition_1 = (resultsinDataFrame['Product-1'] == item_1) & (resultsinDataFrame['Product-2'] == item_2)
    
    # Search for the record where item_2 is in 'Product-1' and item_1 is in 'Product-2'
    condition_2 = (resultsinDataFrame['Product-1'] == item_2) & (resultsinDataFrame['Product-2'] == item_1)
    
    # Combine the conditions using logical OR (|) to find matching records
    matching_record = resultsinDataFrame[condition_1 | condition_2]
    
    # If a matching record is found, return its support value; otherwise, return 0
    if not matching_record.empty:
        return matching_record.iloc[0]['Support']
    else:
        return 0


# Example usage:
columns = ['Input_product', 'Consequent', 'support']
result = pd.DataFrame(columns=columns)


# Display the DataFrame

index = 0
for j in unique_items :
    support_value = search_support(resultsinDataFrame, product, j)
    if support_value != 0 :
        result.loc[index] = [product, j, support_value]
        # print(f"Support for {product} and {j}: {support_value}")
        index =index+1

datafr_sorted = result.sort_values(by='support' ,ascending=False)


plt.rcParams['figure.figsize']=(14,8)
datafr_sorted.plot.bar('Consequent','support',color='Orange')
plt.xlabel('Item Name',fontsize=15)
plt.ylabel('Support Count',fontsize=15)
plt.title('Support of item y with '+product,fontsize=20)
plt.xticks(fontsize=10)
plt.subplots_adjust(left=0.124, bottom=0.276, right=0.943, top=0.9, wspace=0.408, hspace=0.406)
plt.yticks(fontsize=10)
plt.savefig(product+'.png')
plt.show()

# Specify the file path where you want to save the Excel file
file_path = r'C:\Users\bombo\OneDrive\Desktop\eclatresult.xlsx'

# Use pandas to write the DataFrame to an Excel file
datafr_sorted.to_excel(file_path, index=False)

