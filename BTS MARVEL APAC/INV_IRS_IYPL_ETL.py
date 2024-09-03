import os
import glob,time,sys,numpy as np
import pandas as pd,psycopg2 as ps,re
from sqlalchemy import create_engine
from dateutil.relativedelta import relativedelta
from datetime import date, datetime, timedelta
import pytz,re

# Establishing db connection
conn = ps.connect(
    host="digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com",
    database="dgtldbdev",
    user="dgtldevdb",
    password="ahjhskankjs")

conn.autocommit=True
cur=conn.cursor()
# cur.execute(""" delete from bts.bts_inventory_stock bis
# where as_on_date =current_date and  so_known_name  in ('IRS','IYPL') and bts_month=to_char(current_date,'Mon-YY') """)


filename = r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\BTS\ClosingStockRegister_systemfiles\Physical Opening Stock Register.xlsx'
inventory_df = pd.read_excel(filename)  
inventory_df=inventory_df.iloc[1:,:]
filtered_inventory = inventory_df[inventory_df['SITE NAME'].isin(['IRS', 'IYPL'])]
filtered_inventory = filtered_inventory.drop(['SITE ID', 'PARTICULARS', 'DOM/EXP'],axis=1)
filtered_inventory = filtered_inventory.rename(columns={'MONTH': 'bts_month', 'BG/FG': 'product_grade', 'SITE NAME': 'so_known_name',
                                                        'QTY':'quantity_total_stock_in_mt'})

filtered_inventory['as_on_date']=date.today()
filtered_inventory['update_datetime']=datetime.now(pytz.timezone('Asia/Calcutta'))
filtered_inventory['vertical']='PET'
filtered_inventory['company_segment']='IPET'
filtered_inventory['region']='APAC'
filtered_inventory['sales_organization']=''
filtered_inventory['sales_organization_desc']=''
filtered_inventory['plant']=''
# filtered_inventory['plant_desc']=''
# filtered_inventory['material']=''
# filtered_inventory['material_desc']=''
# filtered_inventory['plant_code']=''
# filtered_inventory['product_desc'] = ''
# filtered_inventory['product_bag'] = ''
# filtered_inventory['product_family'] = ''
filtered_inventory['quantity_total_stock_in_kg']=filtered_inventory['quantity_total_stock_in_mt'].apply(lambda x: x*1000)
filtered_inventory['quantity_total_stock_in_lb']=filtered_inventory['quantity_total_stock_in_mt'].apply(lambda x: x*2204.6)
# filtered_inventory['product_desc']=filtered_inventory['material_desc']
filtered_inventory.drop(['Unnamed: 7'],axis=1,inplace=True)
filtered_inventory['bts_month']=filtered_inventory['bts_month'].map(lambda x: x.strftime('%b-%y'))
# filtered_inventory=filtered_inventory.replace('x',np.nan)
quantity_columns = ['quantity_total_stock_in_kg', 'quantity_total_stock_in_lb', 'quantity_total_stock_in_mt']  
for col in quantity_columns:
    filtered_inventory[col] = filtered_inventory[col].fillna(0)

inventory=pd.read_sql(sql=''' select distinct sales_org_code ,sales_org_desc,plant_code,plant_desc,so_known_name 
from bts.plant_company_sales_org_master ''',con=conn)

filtered_inventory.reset_index(inplace=True,drop=True)

for i in range(len(filtered_inventory)):
    for j in range(len(inventory)):
        if filtered_inventory.loc[i]['so_known_name']==inventory.loc[j]['so_known_name']:
            filtered_inventory.at[i,'sales_organization']=inventory.loc[j]['sales_org_code']
            filtered_inventory.at[i,'sales_organization_desc']=inventory.loc[j]['sales_org_desc']
            filtered_inventory.at[i,'plant']=inventory.loc[j]['plant_code']
            filtered_inventory.at[i,'plant_desc'] = inventory.loc[j]['plant_desc']


sap_prod_master = pd.read_sql(sql=''' select distinct so_known_name,material,material_desc,product_name,product_group from bts.sap_product_masters ''',con=conn)
sap_prod_master.reset_index(inplace=True,drop=True)
filtered_inventory.reset_index(inplace=True,drop=True)

# for i in range(len(filtered_inventory)):
#     for j in range(len(sap_prod_master)):
#         if filtered_inventory.loc[i]['so_known_name']==sap_prod_master.loc[j]['so_known_name']:
#             # filtered_inventory.at[i,'material']=sap_prod_master.loc[j]['material']
#             # filtered_inventory.at[i,'material_desc']=sap_prod_master.loc[j]['material_desc']
#             # filtered_inventory.at[i,'plant_code']=sap_prod_master.loc[j]['plant_name']
#             # filtered_inventory.at[i,'product_family']=sap_prod_master.loc[j]['plant_group']
#             # filtered_inventory.at[i,'product_code']=sap_prod_master.loc[j]['product_name']
#             # filtered_inventory.at[i,'product_family']=sap_prod_master.loc[j]['product_group']

cur.execute(""" delete from bts.bts_inventory_stock bis
where as_on_date =current_date and  so_known_name  in ('IRS','IYPL') and bts_month=to_char(current_date,'Mon-YY') """)

pl=','.join(['%s']*len(filtered_inventory.columns))

for idx,row in filtered_inventory.iterrows():

    cur.execute(f"""insert into bts.bts_inventory_stock
(so_known_name,bts_month,product_grade,
quantity_total_stock_in_mt,as_on_date,update_datetime,
vertical,company_segment,region,sales_organization,
sales_organization_desc,plant,quantity_total_stock_in_kg,
quantity_total_stock_in_lb,plant_desc)
values ({pl})""",(row))
    print(idx)
    conn.commit()





