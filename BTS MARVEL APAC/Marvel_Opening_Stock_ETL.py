import pandas as pd
import psycopg2 as ps
from datetime import date,datetime
import pytz,re
import numpy as np

#Establishing Connection
conn=ps.connect(database='dgtldbdev', \
                          host='digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com', \
                              port=5432,user='dgtldevdb',password='Indorama01')
conn.autocommit=True
cur=conn.cursor()

cur.execute(""" delete from bts.bts_inventory_stock where as_on_date =current_date """)


op_df=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\Inventory\Opening_Stock.xlsx',header=1,sheet_name='Sheet1')

cols=[]
_=[cols.append(i) for i in op_df.columns[:4]]
cols.append(op_df.columns[6])
_=[cols.append(i) for i in op_df.columns[16:19]]
op_df=op_df[cols]
op_df.columns=['plant','plant_desc','material','material_desc','bts_month','quantity_total_stock_in_kg',
               'quantity_total_stock_in_lb','quantity_total_stock_in_mt']


master=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx')
op_df['material']=op_df['material'].apply(lambda x : str(x).rsplit('.0')[0] if x!=np.nan else x)
op_df['bts_month']=date.today().strftime('%b-%y')
op_df=op_df[op_df['material'].isin(master['Material'].astype(str).unique())]

op_df['product_desc']=op_df['material_desc']
op_df['product_code']=op_df['material_desc'].apply(lambda x:x.split(',')[1] if len(x.split(','))>1 else None)
op_df['product_quality']=op_df['material_desc'].apply(lambda x:x.split(',')[2] if len(x.split(','))>2 else None)
op_df['product_bag']=op_df['material_desc'].apply(lambda x:x.split(',')[3] if len(x.split(','))>3 else None)
op_df['as_on_date']=date.today()
op_df['update_datetime']=datetime.now(pytz.timezone('Asia/Calcutta'))
op_df['vertical']='PET'
op_df['company_segment']='IPET'
op_df['region']='APAC'
op_df['sales_organization']=''
op_df['sales_organization_desc']=''
op_df['so_known_name']=''
# op_df=op_df[op_df['product_quality']=='A']
op_df.drop('product_quality',axis=1,inplace=True)

master = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx',sheet_name='Combined Mat')

op_df['product_grade']=''
op_df = op_df[op_df['material'].astype(str).isin(master['Material'].astype(str))]
op_df.reset_index(drop=True, inplace=True)
    
for i in range(len(op_df)):
    for j in range(len(master)):
        if op_df.loc[i]['material']==master.loc[j]['Material'] :
            op_df.at[i,'product_grade']=master.loc[j]['Product Group']


sales=pd.read_sql(sql=''' select distinct sales_org_code ,sales_org_desc,plant_code,so_known_name 
from bts.plant_company_sales_org_master ''',con=conn)
op_df.reset_index(inplace=True,drop=True)
for i in range(len(op_df)):
    for j in range(len(sales)):
        if op_df.loc[i]['plant']==sales.loc[j]['plant_code']:
            op_df.at[i,'sales_organization']=sales.loc[j]['sales_org_code']
            op_df.at[i,'sales_organization_desc']=sales.loc[j]['sales_org_desc']
            op_df.at[i,'so_known_name']=sales.loc[j]['so_known_name']

op_df.reset_index(inplace=True,drop=True)

master = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx',sheet_name='Combined Mat')

op_df = op_df[op_df['material'].isin(master['Material'].astype(str))]
op_df.to_csv(r'C:\Users\komalkumari.b\Downloads\opening_stock_output.csv',index=False)


op_df['as_on_date']=op_df['as_on_date'].astype(str)
op_df['update_datetime']=op_df['update_datetime'].astype(str)

pl=','.join(['%s']*len(op_df.columns))

for idx,row in op_df.iterrows():

    cur.execute(f"""insert into bts.bts_inventory_stock
(plant,plant_desc ,material ,material_desc 
,bts_month ,quantity_total_stock_in_kg ,quantity_total_stock_in_lb ,quantity_total_stock_in_mt,product_desc,product_code ,
product_bag,as_on_date,update_datetime,vertical,company_segment,region,sales_organization,
sales_organization_desc,so_known_name,product_grade)
values ({pl})""",(row))
    print(idx)

cur.execute(""" update bts.bts_inventory_stock a set product_grade=null
where product_grade ='' ;
     update bts.bts_inventory_stock a set product_grade =b.bg_fg ,product_family=b.product_group
from bts.sap_product_masters b
where a.material =b.material
and a.as_on_date =current_date ;
            
update bts.bts_inventory_stock  set order_type_desc='EXPORT'
where SPLIT_PART(material_desc , ',', ARRAY_LENGTH(STRING_TO_ARRAY(material_desc, ','),1))='BD';
 
update bts.bts_inventory_stock set order_type_desc='DOMESTIC'
where SPLIT_PART(material_desc , ',', ARRAY_LENGTH(STRING_TO_ARRAY(material_desc, ','),1))='NB' """)




