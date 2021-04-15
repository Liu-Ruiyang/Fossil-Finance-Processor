#!/usr/bin/env python
# coding: utf-8

# In[7]:


import os
import pandas as pd
import numpy as np
import natsort
pd.set_option('display.max_rows', None)
from decimal import *
getcontext().prec = 2

from tkinter  import * 
import tkinter as tk
import tkinter.messagebox as messagebox
from functools import partial


# In[2]:


import zipfile  
def un_zip(file_name):  
    """unzip zip file"""  
    zip_file = zipfile.ZipFile(file_name)  
    if os.path.isdir(file_name + "_files"):  
        pass  
    else:  
        os.mkdir(file_name + "_files")  
    for names in zip_file.namelist():  
        zip_file.extract(names,file_name + "_files/")  
    zip_file.close()  


# In[3]:


def processing(E1,E2,E3):
    try:
        ali_files=E1.get()
        cle_file=E2.get()
        report_path =E3.get()
        
        cle_df = pre_process_cle(cle_file)
        raw_ali_df = merge_ali(ali_files)
        ali_df = pre_process_ali(raw_ali_df)

        report_df = CLE_to_ali(cle_df,ali_df)
        
        report_file_name = report_path + 'Report.csv'
        report_df.to_csv(report_file_name, sep=',', header=True, index=True,encoding='utf_8_sig')
        messagebox.showinfo('Congras', 'Done!')
    except ValueError:
        messagebox.showinfo('Warning', 'Wrong Input Path!')
    


# In[8]:


def create_app():
    window = tk.Tk()

    L1 = Label(window, text="Customer Matcher version 0.1",).grid(row=0,column=1)
    L2 = Label(window, text="Input path of Ali files",).grid(row=1,column=0)
    L3 = Label(window, text="Input name of CLE",).grid(row=2,column=0)
    L4 = Label(window, text="Input path of the report file",).grid(row=3,column=0)
    
    E1 = Entry(window, bd =5)
    E1.grid(row=1,column=1)
    
    E2 = Entry(window, bd =5)
    E2.grid(row=2,column=1)
    
    E3 = Entry(window, bd =5)
    E3.grid(row=3,column=1)
    

    B=Button(window, text ="Submit",command=partial(processing, E1,E2,E3)).grid(row=5,column=1,)
    window.mainloop()


# In[5]:


'''
    path is for file dic
'''
def un_zip_all(path):
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            file_name = os.path.join(root, name)
            if (file_name.endswith('.zip')):
                print(file_name)
#                 un_zip(file_name)

# print(un_zip_all('../Desktop/fossil2020'))


# In[6]:


'''
    Ali files are scv files.
    return a dataframe.
'''
def read_ali(file_name):
    file = pd.read_csv(open(file_name),skiprows=4,header=0,index_col=0,
        names=['账务流水号','业务流水号','商户订单号','商品名称','发生时间','对方账号','收入金额',
      '支出金额（-元）','账户余额（元）','交易渠道',
      '业务类型','备注','业务描述','业务账单来源','业务基础订单号','业务订单号'])
    # drop all data when ID is nan

    file = file.dropna(subset = ['业务基础订单号'])
    df = pd.DataFrame(file)
    df['业务基础订单号'] = df['业务基础订单号'].str.strip('\t')
    df_short = df[['业务基础订单号','收入金额','支出金额（-元）','业务类型']]
    return df_short


# In[7]:


'''
    CLE files are xlsx files.
    return a dataframe.
'''
def read_CLE(file_name):
    file = pd.read_excel(file_name)
    file = file.dropna(subset = ['External Document No.'])
    df = pd.DataFrame(file)
    df_short = df[['External Document No.','Posting Date','Document Type','Customer No.','Amount','Open','User ID']]
    print(file_name + 'Done!')
    return df_short


# In[8]:


'''
    Before use CLE data, need to count overall price for each ID.
    input file and use numpy array;
    sort by IDs; use loop to check and count; delete other price, only remain
    the main ID.
'''
def pre_process_cle(file_name):
    file = pd.read_excel(file_name)
    file = file.dropna(subset = ['External Document No.'])
    df = pd.DataFrame(file)
    df = df[['External Document No.','Posting Date','Document Type','Customer No.','Amount','Open','User ID']]
    
    df_duplicated = df[df['External Document No.'].duplicated()]
    print("Found duplicated error.")
    print(df_duplicated)
    df = df[~df['External Document No.'].duplicated()]
    df = df.sort_values(by='External Document No.')
    
    array = np.array(df)
    new_cle_df = pd.DataFrame(columns=['External Doc No.','Posting Date','Doc Type',
        'Customer No.','Amount','Open','User ID',])
    i = 0
    while i < len(array):
        # Now only check official SUDO127, so pass other customer
        if(array[i,3]!='SUD0127'):
            i = i+1
            continue
        orderid = str(array[i,0])
        if True :
            # this is a header
            head_pos = i
            count = []
            count.append(array[head_pos,4])
        
            for follower_pos in range(head_pos+1,len(array)):
                follower_id = str(array[follower_pos,0])
                # print('follower_id is: '+str(follower_id))
                if orderid in follower_id:
                    count.append(array[follower_pos,4])
                else:
                    break
        
            # count overall value
            result_amount = 0.0
            
            for v in range(0,len(count)):
                # print(str(v)+' : '+str(count[v]))
                result_amount = result_amount+count[v]
            # save 2 .00
            result_amount = round(result_amount, 2)
            # print('now result value is : '+str(result_amount))
            # save new data 
            new_data = {
                            'External Doc No.' : str(orderid),
                            'Posting Date' : array[i,1],
                            'Doc Type' : array[i,2],
                            'Customer No.' : array[i,3],
                            'Amount' : result_amount,
                            'Open' : array[i,5],
                            'User ID' : array[i,6]

                        }
            new_cle_df = new_cle_df.append(new_data,ignore_index=True)
        
    # update i
#             print('now len of count is : '+str(len(count)))
            i = i+len(count)

#     new_cle_df.to_excel(new_cle_file_path, header=True, index=True,encoding='utf_8_sig')
    print('Processing '+file_name+' done!')
    return new_cle_df
    


# In[9]:


'''
    After call function merge_ali,
    need to confine ali_df;
    For single ali_id, When ali type are '交易付款' or '在线支付' or '交易退款'
    Count these three values into one result_value.
    return a dataframe.
'''
def pre_process_ali(ali_df):
    ali_array = np.array(ali_df)
    ali_df_merge = pd.DataFrame(columns=['业务基础订单号','收入金额','支出金额（-元）','业务类型'])
    
    while(len(ali_array!=0)):
        ali_ID = ali_array[0,0]
        # check whether this id has already in ali_df_merge
        temp_ali = np.array(ali_df_merge)
        check = np.argwhere(temp_ali==ali_ID)
        
        if(len(check)==0):
            # print('Now dealing with '+ali_ID)
            locs = np.argwhere(ali_array==ali_ID)
            count = 0.0
            del_index = []
            for loc in locs:
                index = loc[0]
                del_index.append(index)
                ali_income = ali_array[index,1]
                ali_outcome = ali_array[index,2]
                ali_type = ali_array[index,3]
                if(ali_type=='交易付款' or ali_type=='在线支付' or ali_type == '交易退款'):
                    # this is a value needed to be count
                    count = count + ali_income + ali_outcome
                    
                else:
                    continue
            count = round(count,2)
            
            # save this ID and new counted price
            if count >= 0:
                new_data = {
                        '业务基础订单号':ali_ID,
                        '收入金额' : count,
                        '支出金额（-元）' : 0.0,
                        '业务类型' : 'Counted Price'
                }
                ali_df_merge = ali_df_merge.append(new_data,ignore_index=True)
            else:
                new_data = {
                        '业务基础订单号':ali_ID,
                        '收入金额' : 0.0,
                        '支出金额（-元）' : count,
                        '业务类型' : 'Counted Price'
                }
                ali_df_merge = ali_df_merge.append(new_data,ignore_index=True)
            
            # delete related index
#             print(ali_array[0])
            ali_array = np.delete(ali_array, del_index, axis=0)
            

               
        else:
            # this id has been counted before, no need to count, just pass it.
            continue
        # print(ali_ID+" ended.")
    print('Prepocess Ali Done!')
    return ali_df_merge
       


# In[10]:


'''
    Merge all Ali data into one Dataframe;
    path
    return a dataframe.
'''
def merge_ali(path):
    result_df = pd.DataFrame(columns=['业务基础订单号','收入金额','支出金额（-元）','业务类型'])
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            file_name = os.path.join(root, name)
            if (file_name.endswith('.csv') and ('╒╦╬±├≈╧╕' in file_name ) and 
               ('╒╦╬±├≈╧╕(╗π╫▄)' not in file_name)):
                print('Now dealing with '+file_name)
                temp_df = read_ali(file_name)
                result_df = pd.concat([result_df,temp_df],ignore_index = True)
    result_df.set_index(["业务基础订单号"],)
    
    return result_df


# In[11]:


'''
    For efficiency consideration, use numpy to process data;
    Turn dataframe into 2-dimensional numpy;
    
'''
def ali_to_CLE(ali_df,cle_df):
    pass


# In[12]:


'''
    For efficiency consideration, use numpy to process data;
    Turn dataframe into 2-dimensional numpy;
    use the ID in CLE, find the ID in ali, result is a two dimensional list;
    [index][0] , use this to get income and outcome and ready to compare;
    Two bps: 1,no ID in ali; 2, income or outcome cannot match.
    return a dataframe to report.
'''
def CLE_to_ali(cle_df,ali_df):
    cle_array = np.array(cle_df)
    ali_array = np.array(ali_df)
    report_df = pd.DataFrame(columns=['Reason','External Doc No.','Posting Date','Doc Type',
        'Customer No.','Amount','UserID','Open','Ali_Amount','Ali_Type'])
    
    print('Start checking CLE to Ali data.')
    # loop cle_array and search ali_array
    for i in range(0,len(cle_array)):
        print('Now check CLE num'+str(i))
        ID = cle_array[i,0]
        
        # find index location by ID, this loc suppose to be unique
        locs = np.argwhere(ali_array==ID)
        
        
        # if doesn't exist
        if len(locs)==0:
            new_data = {
                'Reason':'Empty',
                'External Doc No.' : str(cle_array[i,0])+'\t',
                'Posting Date' : cle_array[i,1],
                'Doc Type' : cle_array[i,2],
                'Customer No.' : cle_array[i,3],
                'Amount' : cle_array[i,4],
                'Open' : cle_array[i,5]
            }
            report_df = report_df.append(new_data,ignore_index=True)
        else:
            # multiple transaction in one case, need to add 3 types into all;
            report_amount = 0.0
            for loc in locs:
                    
                index = loc[0]
                
                ali_ID = ali_array[index,0]
                if ali_array[index,1]==0 and ali_array[index,2]==0:
                    report_amount = report_amount
                elif ali_array[index,1]==0 and ali_array[index,2]!=0:
                    report_amount = report_amount + ali_array[index,2]
                else:
                    report_amount = report_amount + ali_array[index,1]
                        
            report_amount = round(report_amount, 2)
            # only report priceError cases;
            if(report_amount != cle_array[i,4]):

                new_data = {
                            'Reason':'PriceError',
                            'External Doc No.' : str(cle_array[i,0])+'\t',
                            'Posting Date' : cle_array[i,1],
                            'Doc Type' : cle_array[i,2],
                            'Customer No.' : cle_array[i,3],
                            'Amount' : cle_array[i,4],
                            'Open' : cle_array[i,5],
                            'Ali_Amount' : report_amount,
                            'Ali_Type' : ali_array[index,3]
                }
                report_df = report_df.append(new_data,ignore_index=True)
    return report_df


# In[31]:


'''
    check for single ID in ali pay;
'''

# ali_array = np.array(df_result)
# print(ali_array)
# locs = np.argwhere(ali_array=='1308293858510698275')
# print(locs)
# for loc in locs:
#     index = loc[0]
#     print(ali_array[index,0]+','+str(ali_array[index,1])+','+str(ali_array[index,2])+','+ali_array[index,3])


# In[12]:


def main():
    create_app()
if __name__ == '__main__':
    try:
        main()
        print('Done!')
    except Exception as e:
        print('Error: {}'.format(e))


# In[29]:


# check pre_process_cle is right.
# cle_df = pre_process_cle('../Desktop/fossil2020/Customer Ledger Entries FOSSIL_LINDA.xlsx','./test/8-12/CEL3.xlsx')


# In[30]:


# raw_ali_df = merge_ali('./test/8-12/')
# ali_df = pre_process_ali(raw_ali_df)


# In[31]:


# report_df = CLE_to_ali(cle_df,ali_df)
# report_df.to_csv('./test/8-12/report3.csv', sep=',', header=True, index=True,encoding='utf_8_sig')


# In[9]:





# In[ ]:


'''
../../../Desktop/test/8-12/
../../../Desktop/fossil2020/Customer Ledger Entries FOSSIL_LINDA.xlsx
../../../Desktop/Processor/
'''

