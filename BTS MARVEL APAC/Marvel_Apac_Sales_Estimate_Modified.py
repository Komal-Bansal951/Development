import logging,os
from datetime import date
import sys,pandas as pd

logger=logging.getLogger()
today=date.today()

import psycopg2 as ps
user=os.getlogin()
try:
    conn=ps.connect(database='dgtldbdev',\
    host="digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com",\
        port=5432,user="dgtldevdb",password="Indorama01")
    cursor=conn.cursor()
    conn.autocommit=True
     
    cursor.execute(""" delete from bts.bts_to_sell where as_on_date=current_date """)
    print(f"The number of recortds for to sell deleted is :{cursor.rowcount}")

    cursor.execute(f""" INSERT INTO bts.bts_to_sell
(sales_organization, sales_organization_desc,product_desc, product_code,
product_grade, product_bag, bts_month, to_sell_qty_in_mt, to_sell_qty_in_kg, to_sell_qty_in_lb, as_on_date,
update_datetime, sold_to_customer, product_segment, so_known_name, vertical, company_segment, region,order_type_desc,fob_price)
select
c.sales_org_code as sales_organization,
c.sales_org_desc as  sales_organization_desc,
sq.product_description as product_desc ,
sq.product_code as product_code ,
sq.global_product_code as product_grade ,
null product_bag,
sq.bts_month as bts_month ,
sq.to_sell_qnty as to_sell_qty_mt,
sq.to_sell_qnty*1000 as to_sell_qty_kg,
sq.to_sell_qnty*2204.62 as to_sell_qty_lb,
current_date as_on_date,
current_timestamp  AS update_datetime,
sq.bill_to_customer_name as sold_to_customer,
'PET' as product_segment ,
sq.site_name as so_known_name,
'PET' as vertical,
'IPET' as company_segment,
'APAC' as region,
sq.contract_type as order_type_desc,
sq.fob_price               
from bts.bts_to_sell_qty sq
left join bts.plant_company_sales_org_master  AS c ON sq.site_name  = c.so_known_name
where sq.bts_date=current_date """)
    

    cursor.execute('''update bts.bts_to_sell ts set product_family  = sm.global_product_code 
FROM bts.bts_site_product_mapping sm
where  sm.product_code  = ts.product_code ;
                   
    update bts.bts_to_sell ts set product_grade = pm.bg_fg 
    from bts.sap_product_masters pm
    where pm.product_group  = ts.product_family ;
                   ''')

    df=pd.read_sql(sql=""" select * from bts.bts_to_sell where as_on_date=current_date """,con=conn)
    df.to_excel(rf'C:\Users\{user}\OneDrive - Indorama Ventures PCL\MARVEL APAC\Estimates\sales_estimate_output.xlsx',index=False)
    
except Exception as e:
    print(e)
    logger.error("Error thrown")
    raise e
finally:
    conn.close()