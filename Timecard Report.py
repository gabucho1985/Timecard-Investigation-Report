#!/usr/bin/env python
# coding: utf-8

# In[15]:


#1 good code 
#This is code used for raw data
#After using VBA code for data cleaning purposes I wrote Python code below to calculate hrs worked per day then grouped them 
#by week to get the total hrs worked per week
import os    
import shutil 
import numpy as np
import pandas as pd
import csv
import string
from itertools import zip_longest
import re
from collections import Counter
df =pd.read_csv(r'C:\\Users\\GMacias\Documents\Timecard\Employee_inv_time_report.csv')

def calculate_skew(group):
    group['NewSkew']= group.loc[group.PunchCode=='OUT','Hours'].values[0] -group.loc[group.PunchCode=='INN','Hours'].values[0]
    group['Skew']= group.loc[group.PunchCode=='MAEL','Hours'].values[0]- group.loc[group.PunchCode=='MEAL','Hours'].values[0]
    return group
df=df.groupby(['PunchDate']).apply(calculate_skew)
#df   
df['FinalSkew'] = df['NewSkew'] - df['Skew']
df['Employee Total']=df['FinalSkew']/60
#df 
df[['Employee Total']] = df[['Employee Total']].round(2).where(df[['Employee Total']].apply(lambda x: x != x.shift()), '')
df = df.dropna(subset=['Employee Total'])
df.drop(['Skew', 'NewSkew', 'FinalSkew'], axis=1, inplace=True)
#df = df[np.isfinite(df['Employee Total'])]
#df[np.isfinite(df['Employee Total'])]

# g=df.groupby('PunchDate').nth(0)
# h=g.reset_index(drop=True, inplace=True)
# df=df.groupby('PunchDate').nth(0)
# h=g.reset_index(drop=True, inplace=True)


df=df.groupby('PunchDate').first()


# df


df=df.sort_values(by='In', ascending=True)

df=df.drop(['In'], axis=1)
df['index_col'] = df.index


df.reset_index(drop=False, inplace=True)
df['Employee Total'] = pd.to_numeric(df['Employee Total'], downcast='float')
columns_titles = ["Weekday","PunchCode","PunchDate","Hours","Employee Total"]
df=df.reindex(columns=columns_titles)
#df.round(2)
df['Total for week acc to Employee'] = df.groupby(df.index // 6)['Employee Total'].transform('sum')[lambda x: ~(x.duplicated(keep='last'))]

df.drop(['PunchCode', 'Hours'], axis=1, inplace=True)

#df=df['Total for week acc to Employee'].round(2)
# #df.drop(columns=['PunchCode','Hours'])
file_path = r'C:\\Users\\GMacias\Documents\Timecard\full_total_hours_from_employee.csv'
os.remove(file_path)

df.to_csv('full_total_hours_from_employee.csv',header = True)

source = r'C:\Users\GMacias\full_total_hours_from_employee.csv'
destination = r'C:\Users\GMacias\Documents\Timecard'
dest = shutil.move(source, destination) 


# In[16]:


#Code used based on timecards nngr used to pay employee
#Here we grouped data by week in order to get the total hrs worked per week according to the manager
import os    
import shutil 
import numpy as np
import pandas as pd
import csv
import string
from itertools import zip_longest
import re
#import datetime
df =pd.read_csv(r'C:\\Users\\GMacias\Documents\Timecard\investors_data.csv')


df['Out time'] = pd.to_datetime(df['Out time'])
#df.groupby("Out time")["Hours",'Pay Code'].sum()
#df = df[df.groupby("Out time")["Hours"].transform('min')]
#df.groupby(["Out time","Pay Code"],as_index=False).sum()

df=df.groupby('Out time').agg({'Hours':'sum', 'Pay Code':'last', 'Notes':'last', 'Weekday':'last'})
df['index_col'] = df.index
df.reset_index(drop=False, inplace=True)
df=df.sort_values(by='Out time', ascending=False)

df=df.drop(['index_col'], axis=1)
df['Total for week'] = df.groupby(df.index // 6)['Hours'].transform('sum')[lambda x: ~(x.duplicated(keep='last'))]
df = df.replace(np.nan, '', regex=True)
df.tail()
file_path = r'C:\\Users\\GMacias\Documents\Timecard\mynewfilefullyreloaded.csv'
os.remove(file_path)

df.to_csv('mynewfilefullyreloaded.csv',header = True)

source = r'C:\Users\GMacias\mynewfilefullyreloaded.csv'
destination = r'C:\Users\GMacias\Documents\Timecard'
dest = shutil.move(source, destination) 


# In[18]:


#add  VBA code (weekend saturdays and sunday) to final report after running this code 
import os    
import shutil 
import numpy as np
import pandas as pd
import csv
import string
from itertools import zip_longest
import re
#manager
df1 =pd.read_csv(r'C:\\Users\\GMacias\Documents\Timecard\mynewfilefullyreloaded.csv')
#df1
#result_df1 = df1.drop_duplicates(subset=['N/L sample', 'Address', 'Phone Number'], keep='first')
#print(result_df1)
#df1
#employee
df2= pd.read_csv(r'C:\\Users\\GMacias\Documents\Timecard\full_total_hours_from_employee.csv')
result = pd.concat([df2, df1], axis=1, sort=False) #merge columns from different dataframes
result.fillna("")
#df2
#result
result.to_csv('final_freytes2_.csv',header = True)

source = r'C:\Users\GMacias\final_freytes2_.csv'
destination = r'C:\Users\GMacias\Documents\Timecard'
dest = shutil.move(source, destination) 


# In[10]:


#First Code Sort by  DATE!!! not Hours nor PunchCode
#this will give me the weekends which I have to take out manually, store them in another sheet 
#and then include them at the end
import numpy as np
import pandas as pd
import csv
import string
from itertools import zip_longest

from collections import Counter

#from datetime import datetime, timedelta
import re #have to manually copy and paste mynewfile into notebook-5.7.8-py37_0
df =pd.read_csv(r'C:\\Users\\GMacias\Documents\Timecard\Employee_inv_time_report.csv')
#df
# df.sort_values(by=['PunchDate'])
# df.groupby('PunchDate')['Hours'].sum()
# #df
# # #df
# df['PunchDate'] = pd.to_datetime(df['PunchDate'])
# df
# # df['day_of_week'] = df['PunchDate'].dt.dayofweek
# # df = df[df['day_of_week'] > 4]
# # df
# dtype_beforee = type(df['Post Date'])
list1 = df['PunchDate'].tolist()
dtype_afterr = type(list1)
list1= [x for x in list1 if str(x) != 'nan']
# y= len(list1)/7
# print(y)
x=Counter(list1)
#print(x)
for i in x.elements():
    if x[i]<5:
        continue
        print(x)
    
#         continue
    print( "% s : % s" % (i, x[i]), end ="\n")


# In[ ]:


#First Code Sort by  DATE!!! not Hours nor PunchCode
#this will give me the weekends which I have to take out and the include them at the end
import numpy as np
import pandas as pd
import csv
import string
from itertools import zip_longest
#from datetime import datetime, timedelta
import re #have to manually copy and paste mynewfile into notebook-5.7.8-py37_0
df =pd.read_csv(r'C:\\Users\\GMacias\Documents\Timecard\investors_data.csv')
#df
df.sort_values(by=['Out time'])
df.groupby('Out time')['Hours'].sum()

#df
df['Out time'] = pd.to_datetime(df['Out time'])
#df
df['day_of_week'] = df['Out time'].dt.dayofweek
df = df[df['day_of_week'] > 4]
df


