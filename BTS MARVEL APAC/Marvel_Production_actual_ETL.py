import pandas as pd 
import numpy as np 
from datetime import date,datetime
import pytz,time,re,glob
import psycopg2 as ps
from sqlalchemy import create_engine, types


def dataloading(rec,file,cur,table_name):

    start=time.time()
    if table_name=='bts.sap_bts_production_actuals_raw':
        cur.execute(""" delete from bts.sap_bts_production_actuals_raw """)
        query=""" copy {}(material,plant,storage_location,movement_type,    
material_document,posting_date,qty_in_unit_of_entry,  
unit_of_entry,material_description,movement_type_text,
company_code,special_stock,material_docitem,
base_unit_of_measure,quantity) from STDIN DELIMITER ',' CSV HEADER QUOTE '\"' """.format(table_name)
        cur.copy_expert(query,file)
    elif table_name=='bts.bts_production_actual':
        cur.execute(""" delete from bts.bts_production_actual where as_on_date=current_date """)
        query=""" copy {} from STDIN DELIMITER ',' CSV HEADER QUOTE '\"' """.format(table_name)
        cur.copy_expert(query,file)

    query=""" copy {} from STDIN DELIMITER ',' CSV HEADER QUOTE '\"' """.format(table_name)
    cur.copy_expert(query,file)
    cur.execute('''update bts.bts_production_actual set order_type_desc='EXPORT'
                where SPLIT_PART(material_description , ',', ARRAY_LENGTH(STRING_TO_ARRAY(material_description, ','),1))='BD';
                
                update bts.bts_production_actual set order_type_desc='DOMESTIC'
                where SPLIT_PART(material_description , ',', ARRAY_LENGTH(STRING_TO_ARRAY(material_description, ','),1))='NB' ;
                update bts.bts_production_actual a set product_code=b.product_name,product_grade=b.bg_fg,
product_family=b.product_group
from bts.sap_product_masters b
where a.material =b.material::varchar and a.as_on_date =current_date;
                
update bts.bts_production_actual set sales_organization_desc ='PT. IVI – Indonesia'
where so_known_name ='IVIT' ; ''')
                #     if table_name=='bts.sap_bts_production_actuals_raw':
#         # cur.execute('''update bts.sap_bts_production_actuals_raw set "material"  =replace("order" ,'.0','') ''')
#         # cur.execute('''update bts.sap_bts_production_actuals_raw set "vendor"  =replace("vendor" ,'.0','') ''')
#         # cur.execute('''update bts.sap_bts_production_actuals_raw set "reference"  =replace("reference" ,'.0','') ''')
#         # cur.execute('''update bts.sap_bts_production_actuals_raw set "purchase_order"  =replace("purchase_order" ,'.0','') ''')
#     else:
#         cur.execute(""" update bts.bts_production_actual a set product_grade =b.product_group 
# from bts.sap_product_masters b
# where  a.product_code=b.product_name and a.as_on_date =current_date """)
    end=time.time()
    delta=end-start
    print("Time taken for data loading {} records is :{} secs".format(rec,delta))
    

    # conn.close()



if __name__=='__main__':
    conn=ps.connect(database='dgtldbdev', \
                          host='digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com', \
                              port=5432,user='dgtldevdb',password='Indorama01')
    conn.autocommit=True
    cur=conn.cursor()

    final = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\Production\MB51.xlsx',sheet_name='Sheet1')
    final.rename(columns={'Sales Order':'Sales Order1','Sales order item':'Sales order item1'},inplace=True)
    final.columns = final.columns.str.lower()
    final.columns = [re.sub(' ','_',i) for i in final.columns]
    # final.columns = [re.sub(r'[0-9-]+','',i) for i in final.columns]
    final.columns = [re.sub(r'[./-]','',i) for i in final.columns]
    
    # final.columns=['material', 'plant', 'storage_location', 'movement_type',
    #    'special_stock', 'material_document', 'material_docitem',
    #    'posting_date', 'qty_in_unit_of_entry', 'unit_of_entry', 
    #    'material_description', 'name_1', 'movement_type_text', 'asset',
    #    'subnumber', 'counter', 'order', 'routing_number_for_operations',
    #    'document_header_text', 'document_date', 'qty_in_opun',
    #    'order_price_unit', 'order_unit', 'qty_in_order_unit', 'company_code',
    #    'valuation_type', 'batch', 'entry_date', 'time_of_entry',
    #    'amtin_loccur', 'purchase_order', 'smart_number', 'item',
    #    'ext_amount_in_local_currency', 'sales_value', 'reason_for_movement',
    #    'sales_order', 'sales_order_schedule', 'sales_order_item',
    #    'cost_center', 'customer', 'movement_indicator', 'consumption',
    #    'receipt_indicator', 'vendor', 'sales_order1', 'sales_order_item1',
    #    'base_unit_of_measure', 'quantity', 'material_doc_year', 'network',
    #    'activity', 'wbs_element', 'reservation', 'item_number_of_reservation',
    #    'stock_segment', 'debitcredit_ind', 'user_name', 'transevent_type',
    #    'sales_value_inc_vat', 'currency', 'goods_receiptissue_slip',
    #    'item_automatically_created', 'reference', 'original_line_item',
    #    'multiple_account_assignment']

    # final.columns=['material','plant','storage_location','movement_type','material_document',
    #                'posting_date','qty_in_unit_of_entry','unit_of_entry','material_description',
    #                'name','movement_type_text','document_date','order_price_unit','order_unit',
    #                'qty_in_order_unit','company_code','entry_date','amtin_loccur',
    #                'currency']
    
    final.columns=['material', 'plant', 'storage_location', 'movement_type',     
       'material_document', 'posting_date', 'qty_in_unit_of_entry',  
       'unit_of_entry', 'material_description', 'movement_type_text',
       'company_code', 'special_stock', 'material_docitem',
       'base_unit_of_measure', 'quantity']
    
    final.to_csv(r'C:\Users\komalkumari.b\Downloads\BTS_Prod_Actuals.csv',index=False)
    
    rec=len(final)
    print("Data loading started")
    file=open(file=r'C:\Users\komalkumari.b\Downloads\BTS_Prod_Actuals.csv',mode='r',encoding='latin_1')
    dataloading(rec,file,cur,'bts.sap_bts_production_actuals_raw')
    

    # engine = create_engine('postgresql://dgtldevdb:Indorama01@digital-aws-rds-dev-01.celzhyirfdme.ap-south-1.rds.amazonaws.com:5432/dgtldbdev')

    # final.to_sql('sap_bts_production_actuals_raw',schema='bts',con=engine,if_exists='replace',index=False)



#### Main table processing
    final.rename(columns={'qty_in_unit_of_entry':'quantity_in_kg'},inplace=True)
    for col in ('product_quality','line_no','sales_organization','sales_organization_desc','so_known_name','plant_desc'):
        final[col] = ''
    
    sales=pd.read_sql(sql=""" select distinct plant_code,plant_desc,sales_org_code,sales_org_desc,so_known_name  from bts.plant_company_sales_org_master """,con=conn)

    for i in range(len(final)):
        for j in range(len(sales)):
            if final.loc[i]['plant']==sales.loc[j]['plant_code'] :
                final.at[i,'sales_organization']=sales.loc[j]['sales_org_code']
                final.at[i,'sales_organization_desc']=sales.loc[j]['sales_org_desc']
                final.at[i,'so_known_name']=sales.loc[j]['so_known_name']
                final.at[i,'plant_desc']=sales.loc[j]['plant_desc']
    
    final['order_type_desc'] = np.nan
    final=final[final['material_description'].notnull()]
    breakpoint()
    final['product_code']=final['material_description'].apply(lambda x: x.split(',')[1].strip() if len(x.split(','))>1 else '')
    final['product_grade']=''
    final['product_quality']=final['material_description'].apply(lambda x: x.split(',')[2].strip() if len(x.split(','))>2 else '')
    final['product_bag']=final['material_description'].apply(lambda x: x.split(',')[3].strip() if len(x.split(','))>3 else '')
    final['product_desc']=final['material_description']
    
    # master_data = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Unified Product List.xlsx')

    # master_data['Product Name'] = master_data['Product Name'].astype(str)
    # final=final[final['product_code'].isin([i for i in master_data[Product Name'].unique()])]

    master = pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\Masters\Product master list\Combined Mat2.xlsx',sheet_name='Combined Mat')
    
    # final = final['material'].astype(str).str.split('.')[0]
    final['material'] = final['material'].apply(lambda x: str(x).split('.')[0])
    
    final = final[final['material'].astype(str).isin(master['Material'].astype(str))]
    final.reset_index(drop=True, inplace=True)
    
    final['product_family']=''
    for i in range(len(final)):
        for j in range(len(master)):
            if final.loc[i]['material']==master.loc[j]['Material'] :
                final.at[i,'product_family']=master.loc[j]['Product Group']
                final.at[i,'product_grade']=master.loc[j]['bg_fg']

    # final.drop()
    #final=final.loc[final['product_quality']=='A']
    
    final['region']='APAC'
    final['vertical']='PET'
    final['company_segment']='IPET'
    final['as_on_date']=date.today()
    final['update_datetime']=datetime.now(pytz.timezone('Asia/Calcutta'))
    final.rename(columns={'material_description':'material_description',
                        'posting_date':'production_date'},inplace=True)
    
    final['quantity_in_mt']=final['quantity_in_kg'].apply(lambda x: x*0.001)
    final['quantity_in_lb']=final['quantity_in_kg'].apply(lambda x: x*2.2046)
    final['line_no']=np.nan
    final['bts_month']=final['production_date'].apply(lambda x: x.strftime('%b-%y'))
    final=final[['plant','plant_desc','sales_organization','sales_organization_desc','bts_month','material', 'material_description',
                 'product_desc','product_code','product_grade','product_bag',
                'quantity_in_kg','quantity_in_mt','quantity_in_lb','as_on_date','update_datetime','production_date',
                'region','vertical','company_segment','so_known_name','line_no','product_quality',
                'product_family','order_type_desc']]
    # final['bts_month']=final['bts_month'].map(lambda x: datetime.strptime(x,'%b-%y'))
    # final=final[final['bts_month']>=datetime.strptime(date.today().strftime('%b-%y'),'%b-%y')]

    final.to_csv(r'C:\Users\komalkumari.b\downloads\MB51_ETL.CSV',index= False)
# final = pd.read_csv(r'C:\Users\komalkumari.b\downloads\MB51_ETL.CSV')

rec=len(final)
print("Data loading started")
file=open(file=r'C:\Users\komalkumari.b\downloads\MB51_ETL.CSV',mode='r')
dataloading(rec,file,cur,'bts.bts_production_actual')


