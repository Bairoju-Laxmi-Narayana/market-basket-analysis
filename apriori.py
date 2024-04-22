import numpy as np
import pandas as pd
import xlrd
import openpyxl
import matplotlib.pyplot as plt
from mlxtend.frequent_patterns import apriori as ap
from mlxtend.frequent_patterns import association_rules as ar
from wordcloud import WordCloud
from mlxtend.preprocessing import TransactionEncoder

productName = input()
file_path = r"C:\Users\91934\Downloads\archive\groceries.xlsx"
wb = openpyxl.load_workbook(file_path)

# ws =  wb.sheet_by_index(0)
# for i in range(ws.nrows):
#     for i in range (ws.ncols):
#         print(ws.cell_value(i,j),end="\t")
#     print('')

sheet = wb.active  # You can also use wb['SheetName'] if you know the sheet name

# converting into a dataframe
dataset = sheet.values
groc_data = pd.DataFrame(dataset)
#Apriori Algorithm 

encoding = []
        
for i in range(0, 9835):
    encoding.append([str(groc_data.values[i,j]) for j in range(0, 32)])
    
# Encodes database transaction data in form of a Python list of lists into a NumPy array
# conveting it into an numpy array
encoding = np.array(encoding)
te = TransactionEncoder()

groc_data = te.fit(encoding).transform(encoding)

groc_data = pd.DataFrame(groc_data, columns = te.columns_)


# groc_data=groc_data.drop(['None'],axis=1).astype('bool')
print(groc_data)

frequent_items=ap(groc_data, min_support = 0.001, use_colnames = True)
# frequent_items = frequent_items.drop([60,61],axis=0)
most_pop_items=frequent_items.sort_values('support',ascending=False)
# most_pop_items
most_pop_items=most_pop_items.head(15)
#print(most_pop_items)

# plt.rcParams['figure.figsize']=(5,5)
# most_pop_items.plot.bar('itemsets','support',color='Green')
# plt.xlabel('Item Name',fontsize=10)
# plt.ylabel('Support Count',fontsize=10)
# plt.title('Most Popular Items as per Support',fontsize=20)
# plt.xticks(fontsize=10)
# plt.yticks(fontsize=10)
# plt.show()


# plt.rcParams['figure.figsize']=(5,5)
# wordcloud=WordCloud(background_color = 'black', width = 1200,  height = 1200, max_words = 121).generate(str(most_pop_items))
# plt.imshow(wordcloud)
# plt.axis('off')
# plt.title('Most Popular Items', fontsize=10)
# plt.show()


# plt.rcParams['figure.figsize'] = (5,5)
# color = plt.cm.autumn(np.linspace(0, 1, 40))
# groc_data[0].value_counts().head(40).plot.bar(color = color)
# plt.title('First Most popular items', fontsize = 20)
# plt.xticks(rotation = 90 , fontsize = 7)
# plt.yticks(fontsize = 10)
# plt.show()

association_confi=ar(frequent_items,metric='confidence',min_threshold=0.4)
a_confi_top=association_confi.sort_values('confidence',ascending=False)
print(a_confi_top.drop(['antecedent support','consequent support'],axis=1).head(10))
     

association_supp=ar(frequent_items,metric='support',min_threshold=0.05)
a_supp_top=association_supp.sort_values('support',ascending=False)
print(a_supp_top.drop(['antecedent support','consequent support'],axis=1).head(10))
     

association_lift=ar(frequent_items,metric='lift',min_threshold=3)
a_lift_top=association_lift.sort_values('lift',ascending=False)
print(a_lift_top.drop(['antecedent support','consequent support'],axis=1).head(10))
     
#file = pd.ExcelWriter('C:\\Users\\91934\\Desktop\\confidence.xlsx')
#a_confi_top.drop(['antecedent support','consequent support','leverage','conviction'],axis=1).head(30).to_excel(file)
# assocn_rules_conf['consequents'].to_excel(file)


#changes
dataf = pd.DataFrame(a_lift_top.drop(['antecedent support','consequent support','leverage','conviction'],axis=1).head(30))

# Specify the file path where you want to save the Excel file
file_path = 'C:\\Users\\91934\\Desktop\\confidence.xlsx'

# Use pandas to write the DataFrame to an Excel file
dataf.to_excel(file_path, index=False)

