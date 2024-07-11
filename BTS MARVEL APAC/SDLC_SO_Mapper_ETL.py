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
final=final[final['First Agreed Date']!='#']
final.reset_index(inplace=True,drop=True)
final['First Agreed Date']=pd.to_datetime(final['First Agreed Date'],format='%d.%m.%Y')

final.sort_values(by=['key','First Agreed Date'],ascending=False,inplace=True)
final.drop_duplicates(subset=['key'],inplace=True,ignore_index=True)


final.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\combined_so_delta.xlsx',index=False)

# thai_df=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\SO Delta new.xlsx',sheet_name='so_thailand')
# indo_df=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\SO Delta new.xlsx',sheet_name='so_indonesia')
# df=pd.concat([thai_df,indo_df],axis=0,ignore_index=True)

# df['key']=df[['Sales Order No','Sales document item']].apply(lambda x:str(x[0])+str(x[1]),axis=1)
# df=df[df['First Agreed Date']!='#']
# df.reset_index(inplace=True,drop=True)
# df['First Agreed Date']=pd.to_datetime(df['First Agreed Date'],format='%d.%m.%Y')

# df.sort_values(by=['key','First Agreed Date'],ascending=False,inplace=True)
# df.drop_duplicates(subset=['key'],inplace=True,ignore_index=True)


# df.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\combined_so_delta.xlsx',index=False)


## Bifurcating order B from SDLC

order_b=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC.xlsx',sheet_name='Prompt')

order_b['key']=order_b[['Order Number','Order Item']].apply(lambda x:str(x[0])+str(x[1]),axis=1)
order_b=order_b.loc[order_b['Order Status']=='B']

comb_so=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SO Delta\combined_so_delta.xlsx')


def xlookup(lookup_value):
    lookup_range = comb_so['key']
    return_range = comb_so['SO Outstanding Quantity (MT)']
    df = pd.DataFrame({'lookup': lookup_range, 'return': return_range})
    result = df.loc[df['lookup'] == lookup_value, 'return'].values
    return result[0] if len(result) > 0 else 0

# Example usage
comb_so['key']=comb_so['key'].astype(str)

order_b['outstanding qty in (mt)']=order_b['key'].apply(xlookup)
actual_b=order_b.loc[order_b['outstanding qty in (mt)']<1]
planned_b=order_b.loc[order_b['outstanding qty in (mt)']>=1]

with pd.ExcelWriter(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_OrderStatus_B.xlsx') as writer:
    actual_b.to_excel(writer, sheet_name='actuals', index=False)
    planned_b.to_excel(writer, sheet_name='planned', index=False)

#### 
actual_b['schedule_date']=actual_b.apply(lambda x:x['Requested Delivery Date'] if x['Actual GI Date']=='#' else x['Actual GI Date'],axis=1)
planned_b['schedule_date']=planned_b['Requested Delivery Date']
order_b=pd.concat([actual_b,planned_b],axis=0,ignore_index=True)

order_b.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_OrderStatus_B.xlsx',index=False)

agg=order_b.groupby(['key']).agg({'Order Quantity in (mt)':'mean','Unit Price (USD/kg)':'mean','Confirmed Qty in (mt)':'mean',
                                  'Unconfirmed Qty in (mt)':'mean','Delivery Quantity in (mt)':'sum','outstanding qty in (mt)':'mean'}).reset_index()

mis_col=[col for col in order_b.columns if col not in agg.columns]
for col in mis_col:
    agg[col]=np.nan
for i in range(len(agg)):
    print(i)
    for j in range(len(order_b)):
        if agg.loc[i]['key']==order_b.loc[j]['key']:
            for col in mis_col:
                agg.at[i,col]=order_b.loc[j][col]
            break;

agg=agg[order_b.columns.to_list()]
agg.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_OrderStatus_B.xlsx',index=False)
    
exit()
# order_b.drop_duplicates(subset=['key'],inplace=True,ignore_index=True)
# order_b.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_OrderStatus_B_planned.xlsx',index=False)

##Bifurcating actual and pending orders
ord_b=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_OrderStatus_B.xlsx',index=False)
actual_b=ord_b.loc[ord_b['']<1]
planned_b=ord_b.loc[ord_b['']>=1]
actual_b.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_B_actual.xlsx',index=False)
planned_b.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_B_planned.xlsx',index=False)


#####Combining dispatched and pending B status
actual_b=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_B_actual.xlsx')
actual_b=actual_b[actual_b['Outstanding Qty in (mt)'].isna()] 
actual_b=actual_b[~actual_b['key'].isin([0,'0'])] 


planned_b=pd.read_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_B_planned.xlsx')
planned_b=planned_b[planned_b['Outstanding Qty in (mt)'].notna()] 
planned_b.to_excel(r'C:\Users\komalkumari.b\Downloads\planned_b.xlsx',index=False)

order_b_combined=pd.concat([actual_b,planned_b],axis=0,ignore_index=True)
order_b_combined.to_excel(r'C:\Users\komalkumari.b\OneDrive - Indorama Ventures PCL\MARVEL APAC\SAP Data\SDLC\SDLC_B_Combined.xlsx',index=False)
