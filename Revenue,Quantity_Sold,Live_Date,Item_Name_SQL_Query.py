import pandas as pd
import numpy as np

import sys
import os

from datetime import datetime, timedelta

sys.path.append(os.path.expanduser('~/Analytics/KenMyers/Utils'))
from SQL.Query import Query
from Email.SendMail import SendMail

import matplotlib.pyplot as plt
import seaborn as sns


from IPython.display import display, HTML

display(HTML(data="""
<style>
    div#notebook-container    { width: 95%; }
    div#menubar-container     { width: 65%; }
    div#maintoolbar-container { width: 99%; }
</style>
"""))

today_date = pd.to_datetime('today')
last_year = datetime(today_date.year-1, today_date.month, today_date.day)



#this is the query
q = """
                                                                                                                            
select item_id, sum(sum_of_sales) as Total_Sales, sum(units_sold) as Total_Quantity_sold, live_date::date , name from       
(                                                                                                                           
                                                                                                                            
Select live_items.item_id, case when sum_sales.sum is null then 0 else sum_sales.sum end Sum_of_Sales, 
case when sum_sales.units_sold is null then 0 else sum_sales.units_sold End Units_sold, live_items.live_date

from                                                                                                                        
                                                                                                                            
                                                                                                                            
((select sum(os.price*os.quantity), sum(os.quantity) as Units_sold, left(sku::VARCHAR, 5)  as item_number                   
from orders o                                                                                                               
join order_sku os on o.order_id = os.order_id                                                                               
                                                                                                                            
                                                                                                                            
                                                                                                                            
where o.is_cancelled != 1                                                                                                   
group by  left(sku::VARCHAR, 5))                                                                                            
union ALL                                                                                                                   
 (select sum(oso.price*oso.quantity), sum(oso.quantity) as Units_sold, left(sku::VARCHAR, 5)  as item_number                
from orders_old oo                                                                                                          
join order_sku_old oso on oo.order_id = oso.order_id                                                                        
 where oo.is_cancelled != 1                                                                                                 
group by  left(sku::VARCHAR, 5))) sum_sales                                                                                 
right join                                                                                                                  
    (select i.item_id, i.live_date , i.name from item i                                                                     
where live_date is not null                                                                                                 
and status_id = 1) as Live_items                                                                                            
    on Live_items.item_id = sum_sales.item_number::NUMERIC                                                                  
                                                                                                                            
group by  live_items.item_id, sum_sales.sum, sum_sales.units_sold, live_items.live_date, Live_items.name) all_Sales_Data_For_Live_items
where item_id != 99999                                                                                                      
group by all_Sales_Data_For_Live_Items.item_id, all_Sales_Data_For_Live_Items.live_date, all_Sales_Data_For_Live_Items.name 


"""

# ugpostgres, marketing
data = Query('ugpostgres',q)

data.to_csv("/opt/mnt/publicdrive/Analytics/Gerard/MarketingProjects/ItemVideoPriorities/RevenueAndQuantityData.csv", index=False)