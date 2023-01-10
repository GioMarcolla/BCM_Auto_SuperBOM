# import numpy as np
from pandas import DataFrame as DF
from pandas import read_excel as RE
from pandas import isna
from pandas import concat
# import os
from datetime import date 

DF_FGs = DF(RE("./FGs.XLSX"))
DF_MATs = DF(RE("./MATs.XLSX"))
DF_MRP = DF(RE("./MRP.XLSX"))

DF_FGs.head()
DF_MATs.head()
DF_MRP.head()

DF_FGs.drop(['Material Type', 'Material Group', 'Purchasing Group', 'Vendor', 
             'Name', 'Manufacturer Part No.', 'Manufacturer', 'BCM Stock',
             'Base Unit of Measure', 'ABC Indicator', 'Contract Lead Time',
             'Material Master L-T (Plannned Delv.Time)', 'Total Demand',
             'Purchase Info Record Lead Time'], axis=1, inplace=True)

DF_MATs.drop(['Material Type', 'Material Group', 'Page format', 'Unit',
              'Manufacturer Part No.', 'Manufacturer', 'Planned Deliv. Time',
              'Actual PO Price', 'Actual Currency', 'Last PO Price', 'Currency',
              'VIPA Price', 'Currency', 'Total Req. Qty.', 'Currency.1',
              'Size/dimensions'], axis=1, inplace=True)

DF_MRP.drop(['Material Type', 'Material Group', 'Purchasing Group', 'Vendor', 
             'Name', 'Manufacturer Part No.', 'Manufacturer', 'BCM Stock',
             'Base Unit of Measure', 'ABC Indicator', 'Contract Lead Time',
             'Material Master L-T (Plannned Delv.Time)', 'Total Demand',
             'Purchase Info Record Lead Time',
             'BOM has alternative to this component'], axis=1, inplace=True, errors='ignore')
DF_FGs = DF_FGs[DF_FGs.Description != "PO Item"]
DF_FGs = DF_FGs[DF_FGs.Description != "Balance (sum stock)"]
DF_FGs = DF_FGs[DF_FGs.Description != "PO Confirm"]
DF_FGs = DF_FGs.dropna()

if not 'Customer Material Number' in DF_FGs.columns: DF_FGs.insert(2, 'Customer Material Number', '')
DF_FGs.rename(columns={
    "Sum Demand": "Sum",
    }, inplace=True)


DF_FGs.iloc[:, 4:] *= -1

agg_func = {
    'Material': lambda x: list(x)[0],
    'Material Description': lambda x: ' '.join(set(x)),
    'Customer Material Number': lambda x: ' '.join(set(x)),
    'Description': lambda x: 'Demand'
}

agg_func.update({n:"sum" for n in list(DF_FGs.columns[4:])})

DF_FGs = DF_FGs.groupby(DF_FGs['Material']).aggregate(agg_func)


DF_FGs.head()

nMats = DF_MATs.columns.size - 4
mats_headers = DF_MATs.columns[4:]
DF_MATs.head()

DF_MRP = DF_MRP[DF_MRP.Description != "Safety Stock"]
DF_MRP = DF_MRP[DF_MRP.Description != "OrdRes"]
DF_MRP = DF_MRP[DF_MRP.Description != "Simulation"]
DF_MRP = DF_MRP.dropna()
date_headers = DF_MRP.columns[4:-2]
# print(DF_MRP.head())

DF_Final = DF_MATs.set_index('Material').combine_first(DF_MRP.set_index('Material')).reset_index()
final_headers = [*mats_headers, 'Material', 'Material Description', 'Customer Material Number', 'Description', 'Material Availability', *date_headers, 'Sum Demand', 'Future']

DF_Final = DF_Final.reindex(final_headers, axis=1)

DF_Final.rename(columns={
    "Material Availability": "BCM Stock",
    "Sum Demand": "Sum",
    }, inplace=True)

DF_Final = DF_Final.iloc[:,1:]
DF_Final.replace(to_replace="Balance (sum stock)", value="Stock", inplace=True)
DF_Final.head()
whos = input("Company name: ")
# fn = whos + date.today() + ".xlsx"

DF_Final.to_excel("./" + whos + '_' + str(date.today()) + "_01.xlsx", index=False)
DF_FGs.to_excel("./" + whos + '_' + str(date.today()) + "_02.xlsx", index=False)