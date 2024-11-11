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
    password="xyz")

conn.autocommit=True
cur=conn.cursor()

df_db=pd.read_sql(sql=''' select sales_organization, sales_organization_desc, sap_material, product_desc, sap_material_desc,
product_code, product_grade, bts_month, product_bag, to_sell_qty_in_mt, to_sell_qty_in_kg,
to_sell_qty_in_lb, as_on_date, update_datetime, sold_to_customer, product_segment, so_known_name,
vertical, company_segment, region, order_type_desc, product_family, fob_price
 from bts.bts_to_sell bts where as_on_date = current_date''',con=conn)


user_name=os.getlogin()
# df = pd.read_csv(r'C:\Users\komalkumari.b\Downloads\to_sell_thai_output.csv')
fob= pd.read_excel(rf'C:\Users\{user_name}\OneDrive - Indorama Ventures PCL\BTS\PRICE_Forecast\MeltCost_PriceEstimate_ICIS.xlsx',sheet_name='Price_Forecast',header=1)

# domestic_columns = ['Months', 'Thai China L', 'India']
# df_domestic = df[domestic_columns]

df_dom = fob.iloc[:, :6]

df_dom.rename(columns={'Unnamed: 0':'Months'},inplace=True)
df_dom=df_dom[df_dom['Months']>=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y')]
df_dom['Months']=df_dom['Months'].dt.strftime('%b-%y')

df_dom.set_index('Months', inplace=True)
df_domestic = df_dom[ ['Thai China L','India']]



df_exp = fob.iloc[:, [0] + list(range(6, fob.shape[1]))]
df_exp.rename(columns={'Unnamed: 0':'Months','India.1':'India'},inplace=True)

df_exp.set_index('Months', inplace=True)
df_export = df_exp[['Thai','India']]


thai_sites = ['IRP','IRPL','IPIR']
india_sites = ['ASPET','MPET','IYPL']

# def update_fob_price(df, df_source, sites, fob_column):
#     for site in sites:
#         df_site = df[df['so_known_name'] == site]
#         df_site_merged = df_site.merge(df_source, left_on='bts_month', right_index=True, how='left')
#         df_site[fob_column] = df_site_merged[fob_column]
#         df.update(df_site)
#     return df


def update_fob_price(conn, df_source, sites, fob_column, order_type_desc):
    cur = conn.cursor()
    for month in df_source.index:
        for site in sites:
            fob_price = df_source.at[month, fob_column]
            # month=month.strftime('%b-%y')
            
            update_query = f'''
            UPDATE bts.bts_to_sell
            SET fob_price = {fob_price}
            WHERE bts_month = '{month}' AND so_known_name = '{site}' AND order_type_desc = '{order_type_desc}'
            and as_on_date='2024-06-24'
            '''
            cur.execute(update_query)
            conn.commit()
    
    conn.commit()
    cur.close()

# df_db['fob_price'] = np.nan

# sales_domestic = df_db[df_db['order_type_desc'] == 'DOMESTIC']
# sales_export = df_db[df_db['order_type_desc'] == 'EXPORT']

# sales_domestic = update_fob_price(sales_domestic, df_domestic, thai_sites, 'Thai China L')
# sales_domestic = update_fob_price(sales_domestic, df_domestic, india_sites, 'India')


for col in ('Thai','India'):
    df_export[col]=df_export[col].fillna(0)

for col in ('India','Thai China L'):
    df_domestic[col]=df_domestic[col].fillna(0)



update_fob_price(conn, df_domestic, thai_sites, 'Thai China L', 'DOMESTIC')

update_fob_price(conn, df_domestic, india_sites, 'India', 'DOMESTIC')


# sales_export = update_fob_price(sales_export, df_export, thai_sites, 'Thai')
# sales_export = update_fob_price(sales_export, df_export, india_sites, 'India')

update_fob_price(conn, df_export, thai_sites, 'Thai', 'EXPORT')
update_fob_price(conn, df_export, india_sites, 'India', 'EXPORT')



# sales_updated = pd.concat([sales_domestic, sales_export])

# sales_updated.to_sql('bts_to_sell', con=conn, schema='bts', if_exists='replace', index=False)

# sales_updated.to_excel('updated_data.xlsx', index=False)

print("FOB prices updated successfully.")
