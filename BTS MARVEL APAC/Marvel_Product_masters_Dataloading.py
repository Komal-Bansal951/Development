import pandas as pd 
import numpy as np
from sqlalchemy import create_engine,text

##DB Connection ###

connection_string = 'postgresql://dgtldevdb:Indorama01@digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com:5432/dgtldbdev'
engine=create_engine(connection_string)


# df=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Unified Product List.xlsx')
# df.columns=[i.lower().replace(' ','_') for i in df.columns]

# df=df[~df['product_family'].isin([None,np.nan])]
# table_name='unified_product_masters'

# df.to_sql(table_name,engine,schema='bts',index=False,if_exists='replace')

df=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx')
df.columns=[i.lower().replace(' ','_') for i in df.columns]

table_name='sap_product_masters'

df.to_sql(table_name,engine,schema='bts',index=False,if_exists='replace')
# cur.execute(""" alter table bts.sap_product_masters alter column material type varchar using material::varchar """)


with engine.begin() as cur:
    cur.execute(text(" alter table bts.sap_product_masters alter column material type varchar using material::varchar "))

cur.commit()