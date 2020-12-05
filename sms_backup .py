
# coding: utf-8

# In[102]:


import xml.etree.ElementTree as et
import pandas as pd
import tkinter as tk
from tkinter import filedialog


# In[103]:


import os
from os import listdir
from os.path import isfile, join


# In[104]:


mypath = os.getcwd()
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

for i in onlyfiles:
    if(len(i)==22 and i[0:3]=='sms'):
        file = i


# In[105]:


file = file[0:18] + '.xml'


# In[106]:


tree = et.parse(file)
root = tree.getroot()


# In[107]:


sms = []
for child in root:
    sms.append(child.attrib)


# In[66]:


data = pd.DataFrame(sms)


# In[67]:


data.drop(['contact_name', 'date', 'date_sent','locked','protocol','read','sc_toa','status','sub_id','subject','toa','type'],
         axis = 1, inplace = True) 


# In[68]:


data.rename(columns={'address':'From','body':'Message','readable_date':'Date'},inplace=True)


# In[69]:


start= tk.Tk()

canvas1 = tk.Canvas(start, width = 300, height = 300, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

def exportExcel ():
    global data
    
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    data.to_excel(export_file_path,index = False, header=True)

saveAsButtonExcel = tk.Button(text='Export Excel', command=exportExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=saveAsButtonExcel)

start.mainloop()

