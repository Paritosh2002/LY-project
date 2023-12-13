import pdfplumber
import pandas 
from datetime import datetime
import re

#file  = "C:/Users/paritosh/Downloads/Sai - Tanishq Birla/Tanishq/TANI -AVENUE_AMBI_719__10_KG_.pdf"
#file="C:/Users/paritosh/Downloads/Sai - Tanishq Birla/Tanishq/TANI -AVENUE_760__TURBHE__10Kg.pdf"
#file = "C:/Users/paritosh/Downloads/Cosoha/REVICED BILLNO6487.pdf"
#file = "C:/Users/paritosh/Downloads/Rajehswari rice/RRM_Bill No_966_15-08-23.pdf"
#file = "C:/Users/paritosh/Downloads/Rajehswari rice/RRM_Bill No_977_16-08-23.pdf"
file = "C:/Users/paritosh/Downloads/Cosoha/ANKHI.pdf609cb1b86834247a628960c8_20230801_19_17_53_PM._TCAKBc_pstUWN.pdf"
with pdfplumber.open(file) as pdf:
    first_page = pdf.pages[0]
    layout_text1 = first_page.extract_text(x_tolerance=3, y_tolerance=3, layout=False, x_density=7.25, y_density=13)
layout_text = first_page.extract_text(x_tolerance=3, y_tolerance=3, layout=True, x_density=7.25, y_density=13)
panda1 = pandas.DataFrame(first_page.extract_table(table_settings={}))
column_names2 = panda1.loc[(panda1[0].str.startswith('Sr', na=False)) | (panda1[0].isin(["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]))]

column_names2_clean=column_names2.dropna(axis=1,how='all')
dic = {}
headers = column_names2_clean.iloc[0]
new_df  = pandas.DataFrame(column_names2_clean.values[1:], columns=headers)
header_list = headers.values.flatten().tolist()
print("header_list: {header_list}")
row_list1 = new_df.loc[0, :].values.flatten().tolist()
for i in range(len(new_df.index)):
    item_list = []
    row_array = new_df.loc[i, :].values.flatten().tolist()
    for j in range(len(header_list)):
        item_split = row_array[j].split('\n', 1)[0]
        item_list.append(str(header_list[j]) + ": " + item_split)
    dic[i + 1] = item_list
invoice_no_pattern = r"INVOICE NO\. :(\d+)"
invoice_no_match = re.search(invoice_no_pattern, layout_text)

if invoice_no_match:
    invoice_no = invoice_no_match.group(1)
    print("Matched Invoice Number:", invoice_no)
else:
    print("Invoice number not found.")
buyer_pattern = r"M/s\.(.+?) INVOICE NO"
buyer_match = re.search(buyer_pattern, layout_text)
if buyer_match:
    buyer_name = buyer_match.group(1).strip()
else:
    buyer_name = "Buyer name not found"
    
lines = layout_text1.split('\n')

seller_name = ""

for i, line in enumerate(lines):
    line_lower = line.lower()
    if "tax invoice" in line_lower and i < len(lines) - 1:
        seller_name = lines[i + 1].strip()
        break
    elif "bill of supply" in line_lower and i < len(lines) - 1:
        seller_name = lines[i + 1].strip()
        break
gst_id = None
fssai = None
unique_gst = set()
invoice_dates=[]
final_amount=[]
gst_matches = re.findall(r"\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z\d]{1}[Z]{1}[A-Z\d]{1}", layout_text)
unique_gst.update(gst_matches)
fssai_matches = re.findall(r"\b[0-9]{14}\b", layout_text)
unique_fssai = fssai_matches
date_pattern = r'\b\d{2}/\d{2}/\d{4}\b'
final_date = set()
dates = re.findall(date_pattern, layout_text)

for i in dates:
    if dates.count(i) > 1:
        final_date.add(i)
pattern_total = r"INVOICE TOTAL (?:₹ )?(\d+\.\d{2})"

match = re.search(pattern_total, layout_text)
if match:
    invoice_total = match.group(1)
else:
    invoice_total = "Invoice total not found."
"""x=0
while column_names2_clean[x][4]!="Qty":
    x+=1
qty = column_names2_clean[x][4]"""
if column_names2_clean[6][4]=="Qty":
    qty = column_names2_clean[6][5]
else:
    qty = column_names2_clean[5][4]

print(column_names2_clean)
print(qty)
"""with pdfplumber.open(file) as pdf:
        first_page = pdf.pages[0]
        layout_text1 = first_page.extract_text(x_tolerance=3, y_tolerance=3, layout=False, x_density=7.25, y_density=13)
layout_text = first_page.extract_text(x_tolerance=3, y_tolerance=3, layout=True, x_density=7.25, y_density=13)
panda1 = pandas.DataFrame(first_page.extract_table(table_settings={}))
print(panda1)
column_names2 = pandas.DataFrame()

#column_names2 = panda1.loc[(panda1[0].str.startswith('S.N.', na=False)) | (panda1[0].isin(["1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10."]))]
s = {"1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10."}
for y,x in enumerate(panda1[0]):
        
        if x =="S.N." or x[:2] in s:
                column_names2 = column_names2.append(panda1.iloc[y])
column_names2_clean=column_names2.dropna(axis=1,how='all')
print(column_names2_clean)
newline_count = column_names2.iloc[1, 0].count('\n')
items = newline_count+1
arr=[]
final_list=[]
if newline_count !=0:
        first_col = column_names2.iloc[1, 1].count('\n')
        first_col_list= column_names2_clean.iloc[1, 1].split()
        number_string = ' '.join(map(str, first_col_list))

    # Calculate the length of each part
        part_length = len(number_string) // items

    # Split the string into n parts
        split_list = [number_string[i * part_length: (i + 1) * part_length] for i in range(items)]
        


        for c in range(8):
                if c!=1:
                        
                        li = column_names2_clean.iloc[1, c].split()
                        arr.append(li)
                else:
                        arr.append(split_list)
        temp = []
        
        for v in range(len(arr[0])):
                temp=[]
                for j in range(len(arr)):
                        temp.append(arr[j][v])
                final_list.append(temp)
        print(final_list)
"""