import os
import glob,time,sys,numpy as np
import pandas as pd,psycopg2 as ps,re
from sqlalchemy import create_engine
from dateutil.relativedelta import relativedelta
import win32com.client as win32
from datetime import date, datetime, timedelta
import pytz


#Establishing Connection
conn=ps.connect(database='dgtldbdev', \
                          host='digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com', \
                              port=5432,user='dgtldevdb',password='xyz')
conn.autocommit=True
cur=conn.cursor()

# # cur.execute(""" delete from bts.bts_inventory_stock where as_on_date =current_date """)


def dataloading(rec,file,cur,table_name):

    start=time.time()
    ## Delete thai domestic to sell data
    cur.execute(""" delete  from bts.bts_to_sell bts where as_on_date =current_date 
and so_known_name in ('IRP','IRPL','IPIR') and order_type_desc ='DOMESTIC' """)
    
    ## Data load for to sell thai domestics
    query=""" copy {}(sales_organization,sales_organization_desc,product_code,product_grade,bts_month,
to_sell_qty_in_mt,to_sell_qty_in_kg,to_sell_qty_in_lb,as_on_date,
update_datetime,sold_to_customer,product_segment,so_known_name,
vertical,company_segment,region,order_type_desc,product_family,fob_price) from STDIN DELIMITER ',' CSV HEADER QUOTE '\"' """.format(table_name)
    cur.copy_expert(query,file)


def parse_sheet_name(sheet_name):
    return datetime.strptime(sheet_name, '%b-%y')

def format_sheet_name(date_obj):
    return date_obj.strftime('%b-%y')

def generate_date_range(start_date, months_ahead):
    return [start_date + timedelta(days=30*i) for i in range(months_ahead + 1)]

# start_date = datetime.today()
# months_ahead = 5

# date_range = generate_date_range(start_date, months_ahead)
# ()
# for date in date_range:
#     sheet = format_sheet_name(date)

# sheet = 'Jun-24'
user_name=os.getlogin()
filepath=rf'C:\Users\{user_name}\OneDrive - Indorama Ventures PCL\BTS\Sales_Domestic_Plan\Sale Domestic Plan.xlsx'
xls=pd.ExcelFile(filepath)
final=pd.DataFrame()
for sh in xls.sheet_names:
    
    try:
        if datetime.strptime(sh,'%b-%y')>=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y'):
            temp_df = pd.read_excel(rf'C:\Users\{user_name}\OneDrive - Indorama Ventures PCL\BTS\Sales_Domestic_Plan\Sale Domestic Plan.xlsx',sheet_name=sh,header=5)
            # 
            # col_index = 6
            # df = df[df.columns[col_index:]]

            temp_df['Plant'] = temp_df['Plant'].replace('IPI','IPIR')
            temp_df['Plant'] = temp_df['Plant'].replace('BPC','IRPL')

            temp_df = temp_df[temp_df['Plant'].isin(['IRPL','IPIR','IRP'])]

            temp_df.rename(columns={'Plant':'so_known_name','Local Customer':'sold_to_customer','GRADE':'product_code','To Sell':'to_sell_qty_in_mt'},inplace = True)
            temp_df['bts_month']=sh
            final=pd.concat([final,temp_df],ignore_index=True)    
        
    except Exception as e:
        continue

()
print(" Combined to sell data exported ")

try:
    df=final    
    df.columns = df.columns.str.strip()  
    # df.columns('so_known_name','sold_to_customer','product_code','to_sell_qty_in_mt')
    df['product_family'] = ''
    df['producat_grade'] = ''
    # df['product_code'] = ''

    master = pd.read_excel(rf'C:\Users\{user_name}\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx',sheet_name='Combined Mat')

    df.reset_index(drop=True, inplace=True)
        
    for i in range(len(df)):
        for j in range(len(master)):
            if df.loc[i]['product_code']==master.loc[j]['Product Name'] :
                df.at[i,'product_family']=master.loc[j]['Product Group']
                df.at[i,'product_grade']=master.loc[j]['Product Grade']


    sales=pd.read_sql(sql=''' select distinct sales_org_code ,sales_org_desc,so_known_name 
    from bts.plant_company_sales_org_master ''',con=conn)
    df.reset_index(inplace=True,drop=True)
    for i in range(len(df)):
        for j in range(len(sales)):
            if df.loc[i]['so_known_name']==sales.loc[j]['so_known_name']:
                df.at[i,'sales_organization']=sales.loc[j]['sales_org_code']
                df.at[i,'sales_organization_desc']=sales.loc[j]['sales_org_desc']

    df['to_sell_qty_in_kg']=df['to_sell_qty_in_mt'].apply(lambda x: x*1000)
    df['to_sell_qty_in_lb']=df['to_sell_qty_in_mt'].apply(lambda x: x*2204.6)
    df['as_on_date']=date.today()
    df['update_datetime']=datetime.now(pytz.timezone('Asia/Calcutta'))
    df['vertical']='PET'
    df['company_segment']='IPET'
    df['region']='APAC'
    df['order_type_desc']='DOMESTIC'
    df['product_segment'] = 'PET'
    df['fob_price'] = ''
    

except Exception as e:
        pass

# # df = pd.read_csv(r'C:\Users\{user_name}\Downloads\to_sell_thai_output.csv')
# fob= pd.read_excel(r'C:\Users\{user_name}\OneDrive - Indorama Ventures PCL\BTS\PRICE_Forecast\MeltCost_PriceEstimate_ICIS.xlsx',sheet_name='Price_Forecast',header=1)

# # domestic_columns = ['Months', 'Thai China L', 'India']
# # df_domestic = df[domestic_columns]

# df_dom = fob.iloc[:, :5]
# ()
# df_dom.set_index('Months', inplace=True)
# df_domestic = df_dom[ 'Thai China L']

# df_merged = pd.merge(df, df_domestic, left_on='bts_month', right_on='Months', how='left')
# df_merged['fob_price'] = df_merged['Thai China L']
# df_merged.drop(columns=['Months', 'Thai China L'], inplace=True)

# df_exp = fob.iloc[:, [0] + list(range(5, df.shape[1]))]


# # export_columns = ['Months', 'Thai', 'India']
# # df_export = df[export_columns]


columns_to_keep = [
    'sales_organization', 'sales_organization_desc','product_code', 'product_grade','bts_month',
    'to_sell_qty_in_mt', 'to_sell_qty_in_kg', 'to_sell_qty_in_lb', 'as_on_date',
    'update_datetime', 'sold_to_customer', 'product_segment', 'so_known_name',
    'vertical', 'company_segment', 'region', 'order_type_desc', 'product_family', 'fob_price']

df= df[columns_to_keep]
df = df.reset_index(drop=True)

df['fob_price']=df['fob_price'].fillna(0)
df.to_csv(rf'C:\Users\{user_name}\Downloads\to_sell_thai_output.csv',index=False)

df=pd.read_csv(rf'C:\Users\{user_name}\Downloads\to_sell_thai_output.csv')

df['fob_price']=df['fob_price'].fillna(0)
()
pl=','.join(['%s']*len(df.columns))
df['as_on_date']=df['as_on_date'].astype(str)
df['update_datetime']=df['update_datetime'].astype(str)

# for idx,row in df.iterrows():

#     cur.execute(f"""insert into bts.bts_to_sell
# (sales_organization,sales_organization_desc,product_code,product_grade,bts_month,
# to_sell_qty_in_mt,to_sell_qty_in_kg,to_sell_qty_in_lb,as_on_date,
# update_datetime,sold_to_customer,product_segment,so_known_name,
# vertical,company_segment,region,order_type_desc,product_family,fob_price)
# values ({pl})""",(row))
#     print(idx)
rec=len(df)
table_name='bts.bts_to_sell'
file=open(file=rf'C:\Users\{user_name}\Downloads\to_sell_thai_output.csv',mode='r')
dataloading(rec,file,cur,table_name)
