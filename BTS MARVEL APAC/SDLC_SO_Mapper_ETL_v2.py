import pandas as pd, numpy as np 
from datetime import date,datetime

filepath=r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\SO Delta new.xlsx'
xls=pd.ExcelFile(filepath)
final=pd.DataFrame()
for sh in xls.sheet_names:
    if sh!='_com.sap.ip.bi.xl.hiddensheet':
        
        temp_df=pd.read_excel(filepath,sheet_name=sh)
        final=pd.concat([final,temp_df],ignore_index=True)

final['key']=final[['Sales Order No','Sales document item']].apply(lambda x:str(x[0])+str(x[1]),axis=1)
final=final[final['Confirm Deliv Date']!='#']
final.reset_index(inplace=True,drop=True)
final['Confirm Deliv Date']=pd.to_datetime(final['Confirm Deliv Date'],format='%d.%m.%Y')

final.sort_values(by=['key','Confirm Deliv Date'],ascending=False,inplace=True)
final.drop_duplicates(subset=['key'],inplace=True,ignore_index=True)


final.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\combined_so_delta.xlsx',index=False)
comb_so=final

### Change file location of SDLC
order_b=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC.xlsx',sheet_name='Prompt')

order_b['key']=order_b[['Order Number','Order Item']].apply(lambda x:str(x[0])+str(x[1]),axis=1)
order_b=order_b.loc[order_b['Order Status']=='B']


### Check of delivery dates for b orders in schedule date
def check_delivery_dates(x,y):
    
    if pd.isnull(x):
        return y
    else:
        return x
            

cols=['Requested Delivery Date', 'Actual Delivery Date', 'Actual GI Date']
for col in cols:
    
    order_b[col]=pd.to_datetime(order_b[col],format='%d.%m.%Y',errors='coerce')
 
 
order_b['schedule_date']=order_b.apply(lambda x:check_delivery_dates(x['Actual GI Date'],x['Actual Delivery Date']),axis=1)
order_b['bts_month']=order_b['schedule_date'].dt.strftime('%b-%y')
order_b['schedule_date'].unique()
order_b.to_excel(r'C:\Users\komalkumari.b\Downloads\sample_file.xlsx',index=False)

## Summing dispatch of orders b
def xlookup(lookup_value):
    lookup_range = comb_so['key']
    return_range = comb_so['SO Outstanding Quantity (MT)']
    df = pd.DataFrame({'lookup': lookup_range, 'return': return_range})
    result = df.loc[df['lookup'] == lookup_value, 'return'].values
    return result[0] if len(result) > 0 else 0

# Example usage
comb_so['key']=comb_so['key'].astype(str)

order_b['outstanding qty in (mt)']=order_b['key'].apply(xlookup)

agg=order_b.groupby(['key','bts_month']).agg({'Order Quantity in (mt)':'mean','Unit Price (USD/kg)':'mean','Confirmed Qty in (mt)':'mean',
                                  'Unconfirmed Qty in (mt)':'mean','Delivery Quantity in (mt)':'sum','outstanding qty in (mt)':'mean'}).reset_index()

agg.to_excel(r'C:\Users\komalkumari.b\Downloads\sample_file.xlsx',index=False)

## Ordering orders b based on latest delivery dates
act_orders=order_b
act_orders.sort_values(by=['key','bts_month','schedule_date'],ascending=False,inplace=True)
act_orders.drop_duplicates(subset=['key','bts_month'],inplace=True,ignore_index=True)
act_orders.to_excel(r'C:\Users\komalkumari.b\Downloads\act_orders.xlsx',index=False)

### Mapping summed dispatch with latest del dates

import  numpy as np
mis_col=[col for col in act_orders.columns if col not in agg.columns]
for col in mis_col:
    agg[col]=np.nan
for i in range(len(agg)):
    print(i)
    for j in range(len(act_orders)):
        if ((agg.loc[i]['key']==act_orders.loc[j]['key']) and (agg.loc[i]['bts_month']==act_orders.loc[j]['bts_month'])):
            for col in mis_col:
                agg.at[i,col]=act_orders.loc[j][col]
            break;

agg=agg[act_orders.columns.to_list()]
agg['qty']=agg['Delivery Quantity in (mt)']
agg.to_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_OrderStatus_B_Actuals.xlsx',index=False)
dispatch_orders=agg

### Pending orders for orders b
def xlookup_date(lookup_value):
    lookup_range = comb_so['key']
    return_range = comb_so['Confirm Deliv Date']
    df = pd.DataFrame({'lookup': lookup_range, 'return': return_range})
    result = df.loc[df['lookup'] == lookup_value, 'return'].values
    return result[0] if len(result) > 0 else 0

def xlookup_dispatch(lookup_value):
    lookup_range = comb_so['key']
    return_range = comb_so['Dispatch Quantity in (MT)']
    df = pd.DataFrame({'lookup': lookup_range, 'return': return_range})
    result = df.loc[df['lookup'] == lookup_value, 'return'].values
    return result[0] if len(result) > 0 else 0

# Example usage
comb_so['key']=comb_so['key'].astype(str)
pend_b=order_b
pend_b.drop_duplicates(subset=['key'],inplace=True,ignore_index=True)

pend_b['schedule_date']=pend_b['key'].apply(xlookup_date)
pend_b['Delivery Quantity in (mt)']=pend_b['key'].apply(xlookup_dispatch)
pend_b['qty']=pend_b['outstanding qty in (mt)']
pend_b=pend_b[pend_b['qty']>=1]
pend_b['schedule_date']=pd.to_datetime(pend_b['schedule_date'],errors='coerce')
pend_b['bts_month']=pend_b['schedule_date'].dt.strftime('%b-%y')
pend_b['Order Status']='BP'
pend_b.to_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_OrderStatus_B_Planned.xlsx',index=False)

total_orders=pd.concat([dispatch_orders,pend_b],axis=0,ignore_index=True)
cols=['Sales Order Date','Invoice Created Date']
for col in cols:
    
    total_orders[col]=pd.to_datetime(total_orders[col],format='%d.%m.%Y',errors='coerce')

total_orders.to_excel(r'C:\Users\komalkumari.b\Downloads\SDLC_OrderStatus_B.xlsx',index=False)
