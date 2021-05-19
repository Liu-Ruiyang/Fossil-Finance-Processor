#!/usr/bin/env python
# coding: utf-8

# In[3]:


import os
import pandas as pd
import numpy as np
import natsort
pd.set_option('display.max_rows', None)
from decimal import *
getcontext().prec = 2

import math
from tkinter import *
import tkinter as tk
import tkinter.messagebox as messagebox
from functools import partial


# In[4]:


def create_app():
    window = tk.Tk()

    L1 = Label(
        window,
        text="Username Matcher For Mandy",
    ).grid(row=0, column=1)
    
    L2 = Label(
        window,
        text="Input file path and name of file",
    ).grid(row=1, column=0)
    

    E1 = Entry(window, bd=5)
    E1.grid(row=1, column=1)

    B = Button(window, text="Submit", command=partial(processing, E1)).grid(
                                                          row=5,
                                                          column=1,
                                                      )
    window.mainloop()


# In[5]:


def processing(E1):
    try:
        # actual main()
        file_name = E1.get()
        check_all(file_name,"匹配","小额用户核对结果.xlsx")
        
        messagebox.showinfo('Congras', 'Done!')
    except ValueError:
        messagebox.showinfo('Warning', 'Wrong Input Path!')


# In[6]:


def get_data(file_name, sheet_name):
    file = pd.read_excel(file_name, sheet_name)
    df = pd.DataFrame(file)
    array = np.array(df)
    return array


# In[7]:


def get_subname(combine):
    if combine.find('(') == -1:
        # only contains user, no name needed to be check
        return "NoName"
    else:
        subname = combine.split('(')[0][-1]
        return subname


# In[8]:


def get_subuser(combine):
    if combine.find('(') != -1:
        subuser = combine.split('(')[1].split(')')[0]
    else:
        subuser = combine
    return subuser


# In[9]:


def check_name(name,subname):
    if name[-1] == subname:
        return True
    else:
        return False


# In[10]:


def check_user(user,subuser):
    if subuser.endswith(".com"):
        # user is login with email
        userhead = user[0:3]
        subuserhead = user[0:3]
        if user.endswith(".com"):
            userend = user.split('@')[1]
            subuserend = subuser.split('@')[1]
        else:
            return False
        if userhead == subuserhead and userend == subuserend:
            return True
        else:
            return False
    else:
        # user is login with tele
        userhead = user[0:3]
        subuserhead = user[0:3]
        userend = subuser[-2:-1]
        subuserend = subuser[-2:-1]
        if userhead == subuserhead and userend == subuserend:
            return True
        else:
            return False


# In[11]:


def check_all(file_name, sheet_name,result_file):
    array = get_data(file_name, sheet_name)
    check_result = []

    for i in range(0,len(array)):

        user = array[i,3].strip()
        name = array[i,4]
        combine = array[i,12]
        # combine name only one character and combine tele only first 3 and end 2 or end email;
        subname = get_subname(combine)
        subuser = get_subuser(combine)

        if check_user(user,subuser):
            if check_name(name, subname) == True:
                check_result.append("true")
            else:
                check_result.append("true, but no name.")
        else:
            check_result.append("false")

    
    file = pd.read_excel(file_name, sheet_name)
    df = pd.DataFrame(file)
    df["Check"] = check_result
    df.to_excel(result_file, sheet_name)


# In[70]:


# check_all("EAG小额4.4-5.1 核对.xlsx","匹配","小额用户核对结果.xlsx")


# In[13]:


create_app()


# In[ ]:




