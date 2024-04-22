
import mlxtend
import numpy as np 
import pandas as pd
import openpyxl
from mlxtend.frequent_patterns.fpgrowth import fpgrowth
import matplotlib.pyplot as plt
from wordcloud import WordCloud

product = input()
file_path = r"C:\Users\bombo\OneDrive\Documents\archive\groceries.xlsx"
wb = openpyxl.load_workbook(file_path)

sheet = wb.active  # You can also use wb['SheetName'] if you know the sheet name

# converting into a dataframe
dataset = sheet.values
df = pd.DataFrame(dataset)
df.fillna(0)

# groc_data.show(10)
print(df)      
items = (df[0].unique())

for i in items:
    print(i)


encoded_vals = []
for index, row in df.iterrows():
    labels = {}
    uncommons = list(set(items) - set(row))
    commons = list(set(items).intersection(row))
    for uc in uncommons:
        labels[uc] = 0
    for com in commons:
        labels[com] = 1
    encoded_vals.append(labels)
encoded_vals[0]
encod_df = pd.DataFrame(encoded_vals)

#Encoding values 

print(encod_df)

df.describe()

freq_items = fpgrowth(encod_df , min_support = 0.005 , max_len=2 ,use_colnames = True)
print("freqqq")
print(freq_items)

unique_items = set()
for itemset in freq_items['itemsets']:
    unique_items.update(itemset)

print("Unique items:", unique_items)


most_popular_items=freq_items.sort_values('support',ascending=False)
most_popular_items = most_popular_items.head(15)
print("most popular items")
print(most_popular_items)

#Top 15 most frequent items

most_popular_items.values.tolist()

# plt.rcParams['figure.figsize']=(3,3)
# most_popular_items.plot.bar('itemsets','support',color='Orange')
# plt.xlabel('Item Name',fontsize=15)
# plt.ylabel('Support Count',fontsize=15)
# plt.title('Most Popular Items(as per Support)',fontsize=20)
# plt.xticks(fontsize=10)
# plt.yticks(fontsize=10)
# plt.show()

from mlxtend.frequent_patterns import association_rules
rules = association_rules(freq_items, metric="confidence", min_threshold=0.05)
rules.head(100)

association_confi=association_rules(freq_items,metric='confidence',min_threshold=0.05)
a_confi_top=association_confi.sort_values('confidence',ascending=False)
print("fpgrowthh")
print(a_confi_top.drop(['antecedent support','consequent support'],axis=1).head(20))

# plt.rcParams['figure.figsize'] = (5,5)
# color = plt.cm.autumn(np.linspace(0, 1, 40))
# df[0].value_counts().head(40).plot.bar(color = color)
# plt.title('First Most popular items', fontsize = 20)
# plt.xticks(rotation = 90 , fontsize = 7)
# plt.yticks(fontsize = 10)
# #plt.grid()

# plt.show()

# plt.rcParams['figure.figsize']=(5,5)
# wordcloud=WordCloud(background_color = 'lightgreen', width = 1500, height = 1500, max_words = 121).generate(str(most_popular_items))
# plt.imshow(wordcloud)
# plt.axis('off')
# plt.title('Most Popular Items',fontsize = 12)
# plt.show()

# file = pd.ExcelWriter('Desktop\Rules.xlsx')
# a_confi_top.drop(['antecedent support','consequent support','leverage','conviction'],axis=1).head(100).to_excel(file)
# # assocn_rules_conf['consequents'].to_excel(file)

# # file.save()


       
association_supp=association_rules(freq_items,metric='support',min_threshold=0.05)
a_supp_top=association_supp.sort_values('support',ascending=False)
print(a_supp_top.drop(['antecedent support','consequent support'],axis=1).head())


association_lift=association_rules(freq_items,metric='lift',min_threshold=3)
a_lift_top=association_lift.sort_values('lift',ascending=False)
print(a_lift_top.drop(['antecedent support','consequent support'],axis=1))

datafr = pd.DataFrame(a_lift_top.drop(['antecedent support','consequent support','leverage','conviction'],axis=1).head(100))

# Specify the file path where you want to save the Excel file
file_path = r'C:\Users\bombo\OneDrive\Desktop\rules.xlsx'

# Use pandas to write the DataFrame to an Excel file
datafr.to_excel(file_path, index=False)


def search_support(freq_items_df, item_1, item_2):
    # Concatenate item names into a tuple for easier comparison
    target_items = (item_1, item_2)
    
    # Search for itemsets containing the two specified items
    matching_itemset = freq_items_df[freq_items_df['itemsets'].apply(lambda x: set(target_items).issubset(x))]
    
    # If matching itemset is found, return its support value
    if not matching_itemset.empty:
        return matching_itemset.iloc[0]['support']
    else:
        return 0

# Example usage:






columns = ['Input_product', 'Consequent', 'support']
result = pd.DataFrame(columns=columns)


# Display the DataFrame

index = 0
for j in unique_items :
    support_value = search_support(freq_items, product, j)
    if support_value != 0 :
        result.loc[index] = [product, j, support_value]
        # print(f"Support for {product} and {j}: {support_value}")
        index =index+1

print(result)

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
file_path = r'C:\Users\bombo\OneDrive\Desktop\fpgrowthresult.xlsx'

# Use pandas to write the DataFrame to an Excel file
datafr_sorted.to_excel(file_path, index=False)
