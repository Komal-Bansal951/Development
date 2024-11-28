
import pandas as pd,numpy as np
from datetime import date,datetime
 
## Initializing DB
import psycopg2 as ps
 
 
conn=ps.connect(database='dgtldbdev',\
host="digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com",\
    port=5432,user="dg",password="abjaskahksj")
cursor=conn.cursor()
conn.autocommit=True
 
all_sites=pd.read_sql(sql=""" SELECT distinct so_known_name  FROM bts.plant_company_sales_org_master """,con=conn)
 
### To Sell Reconcilation
def to_sell_check_data(site):
    ref=int(date.today().strftime('%m'))+7
    cursor.execute(f""" select extract(month from (max(to_date(bts_month,'Mon-YY')))) from bts.bts_to_sell bts
                   where as_on_date =current_date
and so_known_name ='{site}' and to_char(to_date( bts_month,'Mon-YY'),'YYYY')='2024'  """)
   
    result=cursor.fetchall()
   
    try:
        m=ref-int(result[0][0])+1
        r=1
    except Exception:
        m=ref-int(date.today().strftime('%m'))
        r=0
   
    for i in range(r,m):
        if r==1:
            i+=int(result[0][0])-int(date.today().strftime('%m'))
 
        cursor.execute(f""" INSERT INTO bts.bts_to_sell
(sales_organization, sales_organization_desc, product_code,
product_grade, bts_month, as_on_date,
update_datetime,so_known_name, vertical, company_segment,
region, order_type_desc, product_family)
select b.sales_org_code,b.sales_org_desc,'N1','BG',to_char(current_date+interval '{i} month','Mon-YY') "bts_month",
current_date ,now() ,b.so_known_name,'PET','IPET','APAC','DOMESTIC','SSP-NBG'
from bts.bts_to_sell a
join bts.plant_company_sales_org_master b on a.so_known_name =b.so_known_name
where  a.so_known_name='{site}' limit 1
 """)
       
 
### Actual Sales Reconcilation
def actual_sales_check_data(site):
    ref=int(date.today().strftime('%m'))+7
    cursor.execute(f""" select extract(month from (max(to_date(bts_month,'Mon-YY')))) from bts.bts_sales_order bts
                   where as_on_date =current_date
and so_known_name ='{site}' and to_char(to_date( bts_month,'Mon-YY'),'YYYY')='2024'  """)
   
    result=cursor.fetchall()
    try:
        m=ref-int(result[0][0])+1
        r=1
    except Exception:
        m=ref-int(date.today().strftime('%m'))+1
        r=0
   
    for i in range(r,m):
        if r==1:
            i+=int(result[0][0])-int(date.today().strftime('%m'))
           
        cursor.execute(f""" INSERT INTO bts.bts_sales_order
(sales_organization, sales_organization_desc, product_code,
product_grade, bts_month, as_on_date,
update_datetime,so_known_name, vertical, company_segment,
region, order_type_desc, product_family)
select b.sales_org_code,b.sales_org_desc,'N1','BG',to_char(current_date+interval '{i} month','Mon-YY') "bts_month",
current_date ,now() ,b.so_known_name,'PET','IPET','APAC','DOMESTIC','SSP-NBG'
from bts.bts_sales_order a
join bts.plant_company_sales_org_master b on a.so_known_name =b.so_known_name
where a.so_known_name='{site}' limit 1
 """)
 
### Prod estimate Reconcilation
def prod_estimate_check_data(site):
    ref=int(date.today().strftime('%m'))+7
    cursor.execute(f""" select extract(month from (max(to_date(bts_month,'Mon-YY')))) from bts.bts_production_estimate bts
                   where as_on_date =current_date
and so_known_name ='{site}' and to_char(to_date( bts_month,'Mon-YY'),'YYYY')='2024' """)
   
    result=cursor.fetchall()
    try:
        m=ref-int(result[0][0])+1
        r=1
    except Exception:
        m=ref-int(date.today().strftime('%m'))+1
        r=0
   
    for i in range(r,m):
        if r==1:
            i+=int(result[0][0])-int(date.today().strftime('%m'))
        cursor.execute(f""" INSERT INTO bts.bts_production_estimate
(sales_organization, sales_organization_desc, product_code,
product_grade, bts_month, as_on_date,
update_datetime,so_known_name, vertical, company_segment,
region, product_family)
select b.sales_org_code,b.sales_org_desc,'N1','BG',to_char(current_date+interval '{i} month','Mon-YY') "bts_month",
current_date ,now() ,b.so_known_name,'PET','IPET','APAC','SSP-NBG'
from bts.bts_production_estimate a
join bts.plant_company_sales_org_master b on a.so_known_name =b.so_known_name
where a.so_known_name='{site}' limit 1
 """)
 
### Prod actuals Reconcilation
def prod_actual_check_data(site):
    ref=int(date.today().strftime('%m'))+7
    cursor.execute(f""" select extract(month from (max(to_date(bts_month,'Mon-YY')))) from bts.bts_production_actual bts
                   where as_on_date =current_date
and so_known_name ='{site}' and to_char(to_date( bts_month,'Mon-YY'),'YYYY')='2024'  """)
   
    result=cursor.fetchall()
    try:
        m=ref-int(result[0][0])+1
        r=1
    except Exception:
        m=ref-int(date.today().strftime('%m'))+1
        r=0
   
    for i in range(r,m):
        if r==1:
            i+=int(result[0][0])-int(date.today().strftime('%m'))
        cursor.execute(f""" INSERT INTO bts.bts_production_actual
(sales_organization, sales_organization_desc, product_code,
product_grade, bts_month, as_on_date,
update_datetime,so_known_name, vertical, company_segment,
region,order_type_desc, product_family)
select b.sales_org_code,b.sales_org_desc,'N1','BG',to_char(current_date+interval '{i} month','Mon-YY') "bts_month",
current_date ,now() ,b.so_known_name,'PET','IPET','APAC','DOMESTIC','SSP-NBG'
from bts.bts_production_actual a
join bts.plant_company_sales_org_master b on a.so_known_name =b.so_known_name
where a.so_known_name='{site}' limit 1
 """)
       
all_sites['so_known_name'].apply(lambda x:to_sell_check_data(x))
all_sites['so_known_name'].apply(lambda x:actual_sales_check_data(x))
all_sites['so_known_name'].apply(lambda x:prod_estimate_check_data(x))
all_sites['so_known_name'].apply(lambda x:prod_actual_check_data(x))
