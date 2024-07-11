import pandas as pd 
import numpy as np 
from datetime import date,datetime,timedelta
import pytz,time,re,glob
import psycopg2 as ps

#Establishing Connection
conn=ps.connect(database='dgtldbdev', \
                          host='digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com', \
                              port=5432,user='dgtldevdb',password='Indorama01')
conn.autocommit=True
cur=conn.cursor()

cur.execute(""" delete from bts.bts_min_stock where to_date(bts_month,'Mon-YY')>(SELECT MAX(to_date(bts_month,'Mon-YY'))
FROM bts.bts_min_stock
WHERE min_stock_qty_in_mt IS NOT null and as_on_date =current_date) and as_on_date =current_date;
            """)

min_stock=pd.read_sql(sql=""" SELECT MAX(to_date(bts_month,'Mon-YY')) "max_bts_month"
FROM bts.bts_min_stock
WHERE min_stock_qty_in_mt IS NOT null and as_on_date =current_date """,con=conn)

ref = min_stock['max_bts_month'].unique()[0]

inc=12-int(ref.strftime('%m'))
print(inc)

for loop_counter in range(1,inc+1):
    init=ref
    init=init.replace(month=int(init.strftime('%m'))+loop_counter).strftime('%b-%y')

    cur.execute(f'''INSERT INTO bts.bts_min_stock (sales_organization, sales_organization_desc, sap_material, sap_material_desc, product_desc,
            product_code, product_grade, product_bag, bts_month, min_stock_qty_in_mt, min_stock_qty_in_kg,
            min_stock_qty_in_lb, as_on_date, update_datetime, plant, plant_desc, vertical, company_segment, 
            region, so_known_name,product_family,order_type_desc)
            SELECT sales_organization, sales_organization_desc, sap_material, sap_material_desc, product_desc,
            product_code, product_grade, product_bag, '{init}' bts_month, min_stock_qty_in_mt, min_stock_qty_in_kg,
            min_stock_qty_in_lb, as_on_date, update_datetime, plant, plant_desc, vertical, company_segment, 
            region, so_known_name,product_family,order_type_desc
            FROM bts.bts_min_stock
            WHERE bts_month = (SELECT TO_CHAR(MAX(TO_DATE(bts_month,'Mon-YY')), 'Mon-YY')
                           FROM bts.bts_min_stock
                           WHERE min_stock_qty_in_mt IS NOT NULL AND as_on_date = current_date) AND as_on_date = current_date''')
    
    













































































































































































































































