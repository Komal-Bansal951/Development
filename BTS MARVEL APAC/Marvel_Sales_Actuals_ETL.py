import os
import glob,time,sys,numpy as np
import pandas as pd,psycopg2 as ps,re
from sqlalchemy import create_engine
from dateutil.relativedelta import relativedelta
import win32com.client as win32
from datetime import date, datetime, timedelta
# from Auto_emailer import Email
# from sdlc_dataloading import Transform


# Establishing db connection
conn = ps.connect(
    host="digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com",
    database="dgtldbdev",
    user="dgtldevdb",
    password="Indorama01")

db =conn.cursor()
db1 = pd.read_sql(''' select * from bts.plant_company_sales_org_master pcsom''',con = conn)
db.execute(""" delete from bts.sap_bts_sales_order_raw """)
conn.commit()

# sales_data = pd.DataFrame()

# sales_data=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\Data Models\Sales Data Model.xlsx',sheet_name='full')

# sales_data = sales_data.loc[(sales_data['Order Status'] == 'A') | (sales_data['Order Status'] == 'C')]

# # dis_df=df.loc[df['order_status']=='C']
# # dis_df=dis_df.loc[dis_df['schedule_date']<=np.datetime64(date.today())]
# # dis_df=dis_df[dis_df['schedule_date']>=np.datetime64(datetime.now().replace(day=1))]
# # pen_df=df.loc[df['order_status']=='A']
# # pen_df=pen_df.loc[pen_df['schedule_date']>np.datetime64(date.today())]

# # sales_data=sales_data.loc[(sales_data['order_status']=='A') & (sales_data['schedule_date']>np.datetime64(date.today()))]
# # sales_data=sales_data.loc[(sales_data['order_status']=='C') & (sales_data['schedule_date']<=np.datetime64(date.today()))]

# sales_B = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\Sales_data_model_temp.xlsx')

# final_df = pd.concat([sales_data,sales_B],axis=0)
# final_df = final_df.drop(columns=['key'])

# final_df.reset_index(drop=True, inplace=True)

# final_df.columns = map(str.lower, final_df.columns)
# final_df.columns = final_df.columns.str.replace(' ', '_')
# rename_cols=[re.sub(r'[.)(/?-]+','',i) for i in final_df.columns]
# rename_cols=[re.sub(r' ','_',i) for i in rename_cols]
        
# final_df.rename(columns={'company_code_name' : 'company_code_desc', 'plant' : 'plant_code' , 'plant_name' : 'plant_desc','sales_organization_name':'sales_org_desc',
#                                                  'qty' : 'sales_qty_mt' , 'division_desc' : 'division_name' ,'ship_to_customer_name':'ship_to',
#                                                  'sold_to_country_desc' : 'sold_to_country_name' , 'sold_to_customer_name' : 'soldto' , 
#                                                  'ship_to_country_desc' :'ship_to_country_name','material__-_sales_order' : 'material___sales_order', 
#                                                  'material__-_sales_order_desc' : 'material_sales_orders_desc' , 'invoice_created_date' : 'sales_invoice_date' , 
#                                                  'order_quantity_in_(mt)' :'order_quantity_in_mt','unit_price_(usd/kg)':'unit_price_usdkg',
#                                                  'confirmed_qty_in_(mt)':'confirmed_qty_in_mt','unconfirmed_qty_in_(mt)':'unconfirmed_qty_in_mt',
#                                                  'delivery_quantity_in_(mt)':'delivery_quantity_in_mt',
#                                                  'billing_quantity_in_(mt)':'billing_quantity_in_mt'},inplace=True)



# final_df['os_qty']=final_df['sales_qty_mt']-final_df['delivery_quantity_in_mt']
# final_df.to_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_logic.xlsx',index=False)


# engine = create_engine('postgresql://dgtldevdb:Indorama01@digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com:5432/dgtldbdev')
 
# final_df=final_df.replace('x',np.nan)
# final_df.to_sql('sap_bts_sales_order_raw',schema='bts',con=engine,if_exists='replace',index=False)


# #### Main table
# final_df=final_df[final_df['bts_month'].notna()]
# breakpoint()
# for i in range (len(final_df)):
#     try:
#         if not isinstance(final_df.loc[i]['bts_month'],str):
#             final_df.at[i,'bts_month'] = final_df.loc[i]['bts_month'].strftime('%b-%y')
#         else: 
#             final_df.at[i,'bts_month'] = final_df.loc[i]['bts_month']
#     except Exception :
#         pass

# # final_df['bts_month']=final_df['bts_month'].apply(lambda x: x.strftime('%b-%y'))
# final_df.to_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_output.xlsx',index=False)

# dis_df=final_df.loc[final_df['order_status']=='C']
# dis_df=dis_df.loc[dis_df['schedule_date']<=np.datetime64(date.today())]
# dis_df=dis_df[dis_df['schedule_date']>=np.datetime64(datetime.now().replace(day=1))]
# pen_df=final_df.loc[final_df['order_status']=='A']
# pen_df=pen_df.loc[pen_df['schedule_date']>=np.datetime64(datetime.now().replace(month=int(date.today().strftime('%m'))-1 or 12,day=1))]

# pen_b=final_df.loc[final_df['order_status']=='B']
# pen_b=pen_b.loc[pen_b['schedule_date']>=np.datetime64(datetime.now().replace(month=int(date.today().strftime('%m'))-1 or 12,day=1))]


# final_df=pd.concat([dis_df,pen_df,pen_b],ignore_index=True)
# final_df.loc[final_df['schedule_date']<np.datetime64(datetime.now().replace(day=1)),'bts_month']=date.today().strftime('%b-%y')
final_df = pd.read_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_output.xlsx')
master = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx',sheet_name='Combined Mat')

final_df = final_df[final_df['material___sales_order'].isin(master['Material'])]
final_df.reset_index(inplace=True)
# master = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx',sheet_name='Combined Mat')
for i in range(len(final_df)):
    for j in range(len(master)):
        if final_df.loc[i]['material___sales_order']==master.loc[j]['Material'] :
            final_df.at[i,'product_grade']=master.loc[j]['Product Group']
            final_df.at[i,'material_sales_orders_desc']=master.loc[j]['Material_Desc']

final_df=pd.read_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_output.xlsx')
final_df.rename(columns={'material___sales_order':'material_sales_order','material_sales_orders_desc':'material_sales_order_desc','product_name':'product_code','unit_price_usdkg':'unit_price_usd_kg',
                   'soldto':'sold_to','sold_to_customer':'sold_to_desc','ship_to_customer':'ship_to_desc',
                   'sales_org_desc':'sales_organization_desc'},inplace=True)
# final_df=final_df.loc[final_df['product_quality']=='A']
final_df=final_df.loc[final_df['material_sales_order_desc']!=0]

final_df['product_bag']=final_df['material_sales_order_desc'].apply(lambda x: x.split(',')[3] if len(x.split(','))>3 else None)
final_df['product_desc']=final_df['material_sales_order_desc']
final_df['planned_gi_date_order']=np.nan
final_df['order_quantity_in_kg']=final_df['order_quantity_in_mt']*1000
final_df['order_quantity_in_lb']=final_df['order_quantity_in_mt']*2204.62
final_df['confirmed_qty_in_kg']=final_df['confirmed_qty_in_mt']*1000
final_df['confirmed_qty_in_lb']=final_df['confirmed_qty_in_mt']*2204.62
final_df['delivery_quantity_in_kg']=final_df['delivery_quantity_in_mt']*1000
final_df['delivery_quantity_in_lb']=final_df['delivery_quantity_in_mt']*2204.62
final_df['sales_qty_kg']=final_df['sales_qty_mt']*1000
final_df['sales_qty_lb']=final_df['sales_qty_mt']*2204.62
final_df['unit_price_dc']=final_df['unit_price_usd_kg']
final_df['document_currency']='USD'
final_df['vertical']='PET'
final_df['company_segment']='IPET'
final_df['region']='APAC'
final_df['as_on_date']=date.today()
final_df['update_datetime']=datetime.now()
final_df['os_qty']=final_df['sales_qty_mt']-final_df['delivery_quantity_in_mt']

# db1 = pd.read_sql(''' select * from bts.plant_company_sales_org_master pcsom''',con = conn)

sales=pd.read_sql(sql=""" select distinct sales_org_code,sales_org_desc,so_known_name from bts.plant_company_sales_org_master """,con=conn)

final_df['so_known_name'] = ''

final_df.reset_index(drop=True, inplace=True)

# for i in range(len(final_df)):
#     for j in range(len(sales)):
#         if final_df.loc[i]['sales_organization']==sales.loc[j]['sales_org_code'] :
#             final_df.at[i,'so_known_name']=sales.loc[j]['so_known_name']
#             final_df[i,'sales_organization_desc'] = sales.loc[j]['sales_org_desc']    

c=int(date.today().strftime('%m'))
y=int(date.today().strftime('%Y'))
if c+5>12:
    c=(c+5)-12
    y=y+1
    b=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y').replace(day=1,month=c,year=y)
else:
    b=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y').replace(day=1,month=c+5,year=y)

final_df['bts_month']=final_df['bts_month'].map(lambda x: datetime.strptime(x,'%b-%y'))
final_df=final_df[final_df['bts_month']>=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y')]
final_df=final_df[final_df['bts_month']<b]

final_df['bts_month']=final_df['bts_month'].map(lambda x: x.strftime('%b-%y'))


final_df=final_df[['sales_organization','sales_organization_desc','order_number','order_item','order_type','order_type_desc',
'material_sales_order','material_sales_order_desc','product_desc','product_code','product_quality',
'product_bag','sales_order_date','actual_gi_date','planned_gi_date_order','bts_month','order_quantity_in_mt',
'order_quantity_in_kg','order_quantity_in_lb','confirmed_qty_in_mt','confirmed_qty_in_kg','confirmed_qty_in_lb',
'delivery_quantity_in_mt','delivery_quantity_in_kg','delivery_quantity_in_lb','unit_price_usd_kg','unit_price_dc',
'document_currency','sales_qty_mt','sales_qty_kg','sales_qty_lb','as_on_date','update_datetime','sold_to',
'sold_to_desc','ship_to','ship_to_desc','vertical','company_segment','region','so_known_name','fob_price','delivery',
'actual_delivery_date','os_qty','product_grade']]

## Create connection using sqlachemy
engine = create_engine('postgresql://dgtldevdb:Indorama01@digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com:5432/dgtldbdev')


db.execute(""" delete from bts.bts_sales_order where as_on_date=current_date """)
conn.commit()

final_df.to_sql('bts_sales_order',schema='bts',con=engine,if_exists='append',index=False)
print('Data loading for SDLC completed')

#sto
user=os.getlogin()
df=pd.read_excel(rf'C:\Users\{user}\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\Data Models\STO Data model.xlsx')
df=df[df['bts_month'].notna()]
 
# df['bts_month']=df['bts_month'].map(lambda x: datetime.strptime(x,'%b-%y'))
df=df[df['bts_month']>=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y')]
df['bts_month']=df['bts_month'].map(lambda x: x.strftime('%b-%y'))
df['as_on_date']=date.today()
## Create connection using sqlachemy 
 
df.to_sql('bts_sales_order',schema='bts',con=engine,if_exists='append',index=False)

# internal consumption 
b=datetime.strptime(datetime.now().replace(day=1).strftime('%d-%m-%Y'),'%d-%m-%Y')
 
df=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\BTS\GIVL\Order_Int_Csmp.xlsx')
df['as_on_date']=date.today()
df['update_datetime']=datetime.now()
df.rename(columns={'actual':'sales_qty_mt'},inplace=True)
df['sales_qty_kg']=df['sales_qty_mt']*1000
df['sales_qty_lb']=df['sales_qty_mt']*2204.6
df['order_quantity_in_mt']=df['sales_qty_mt']
df['order_quantity_in_kg']=df['sales_qty_kg']
df['order_quantity_in_lb']=df['sales_qty_lb']
df['order_type_desc']='INTERNAL_CONSUMPTION'
df['product_grade']='FG'
df['product_family']='AMP-FG'
df['vertical']='PET'
df['company_segment']='IPET'
df['region']='APAC'
df['so_known_name']='GIVL'
indx=[]
for x in range(df.shape[0]):
    if(datetime.strptime(df.loc[x]['bts_month'],'%b-%y')<b):
        indx.append(x)
df.drop(indx,axis=0,inplace=True)
df.reset_index(inplace=True,drop=True)
df['sales_organization']='CN50'
df['sales_organization_desc']='GIVL PET Poly Co.Ltd'
df['as_on_date']=df['as_on_date'].astype('string')
df['update_datetime']=df['update_datetime'].astype('string')
df.drop(['site_id','particulars'],axis=1,inplace=True)
 
for i in range(df.shape[0]):
    lst=list()
    for j in df.columns.to_list():
        lst.append(df.loc[i][j])
    lst=tuple(lst)  
    qry=f"""INSERT INTO bts.bts_sales_order
(bts_month,sales_qty_mt,as_on_date,
update_datetime,sales_qty_kg,sales_qty_lb,order_quantity_in_mt,order_quantity_in_kg,
order_quantity_in_lb,order_type_desc,
product_grade,product_family,vertical,company_segment,
region,so_known_name,sales_organization,
sales_organization_desc) values {lst}"""
    db.execute(qry)                
    print(i)
    del(lst)
 
db.execute("""  update bts.bts_sales_order a set so_known_name =b.so_known_name ,sales_organization_desc =b.sales_org_desc 
from bts.plant_company_sales_org_master b
where a.sales_organization=b.sales_org_code;
           
update bts.bts_sales_order so set order_type_desc =c.contract_type
from (select distinct a.order_number ,b.sold_to_country_name,
case when lower(b.sold_to_country_name)='thailand' then 'DOMESTIC'else 'EXPORT' end as "contract_type"
from bts.bts_sales_order a
join bts.sap_bts_sales_order_raw b
on a.order_number ::varchar=b.order_number::varchar where as_on_date =current_date) c
where sales_organization  = any(array['TH40','TH45','TH10','TH15'])  and
so.order_number = c.order_number ;

update bts.bts_sales_order so set order_type_desc =c.contract_type
from (select distinct a.order_number ,b.sold_to_country_name,
case when lower(b.sold_to_country_name)='indonesia' then 'DOMESTIC'else 'EXPORT' end as "contract_type"
from bts.bts_sales_order a
join bts.sap_bts_sales_order_raw b
on a.order_number ::varchar=b.order_number::varchar where as_on_date =current_date) c
where sales_organization  = any(array['ID20','ID35','IRS'])  and
so.order_number = c.order_number ;

update bts.bts_sales_order so set order_type_desc =c.contract_type
from (select distinct a.order_number ,b.sold_to_country_name,
case when lower(b.sold_to_country_name)='india' then 'DOMESTIC'else 'EXPORT' end as "contract_type"
from bts.bts_sales_order a
join bts.sap_bts_sales_order_raw b
on a.order_number ::varchar=b.order_number::varchar where as_on_date =current_date) c
where sales_organization  = any(array['IYPL','IN40','IN30'])  and
so.order_number = c.order_number ;

update bts.bts_sales_order so set order_type_desc =c.contract_type
from (select distinct a.order_number ,b.sold_to_country_name,
case when lower(b.sold_to_country_name)='china' then 'DOMESTIC'else 'EXPORT' end as "contract_type"
from bts.bts_sales_order a
join bts.sap_bts_sales_order_raw b
on a.order_number ::varchar=b.order_number::varchar where as_on_date =current_date) c
where sales_organization  = any(array['CN50'])  and
so.order_number = c.order_number ;

update bts.bts_sales_order so set order_type_desc =c.contract_type
from (select distinct a.order_number ,b.sold_to_country_name,
case when lower(b.sold_to_country_name)='egypt' then 'DOMESTIC'else 'EXPORT' end as "contract_type"
from bts.bts_sales_order a
join bts.sap_bts_sales_order_raw b
on a.order_number ::varchar=b.order_number::varchar where as_on_date =current_date) c
where sales_organization  = any(array['EG10'])  and
so.order_number = c.order_number;
           
update bts.bts_sales_order a set product_grade =b.bg_fg,product_family =b.product_group
from bts.sap_product_masters b
where a.material_sales_order=b.material and a.as_on_date=current_date;
           
delete from bts.bts_sales_order where product_grade is null
and as_on_date=current_date and order_type_desc !='Intercompany' """)
conn.commit()
conn.close()
# split_values = sales_data['material_sales_orders_desc'].str.split(',', expand=True)
# filtered_df = sales_data[split_values[2] == 'A']
# sales_data[['product_code', 'product_grade', 'product_bag']] = split_values[[1, 2, 3]]
# filtered_df.drop(columns=['material_sales_orders_desc'], inplace=True)

# filtered_data = sales_data[sales_data['product_grade'] == 'A']

# filtered_df = sales_data[sales_data['order_type'].isin(['A', 'C'])]

# sales_data=pd.read_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_logic.xlsx')
# sales_data=sales_data.loc[(sales_data['order_status']=='A') & (sales_data['schedule_date']>np.datetime64(date.today()))]
# sales_data=sales_data.loc[(sales_data['order_status']=='C') & (sales_data['schedule_date']<=np.datetime64(date.today()))]

