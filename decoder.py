#!/usr/bin/env python
# coding: utf-8

# In[141]:


# File import

import pandas as pd

df = pd.read_excel("input.xlsx").fillna("")
df.drop(df.index[0:11], inplace = True)
df.reset_index(inplace = True)
df.drop("index", axis = 1, inplace = True)

print(df.info())
# print(df.head(30))


# In[142]:


# iterate through rows and find Cycle marker

df["new cycle"] = df.iloc[:, 0].str.contains("Cycle")


# print(df["new cycle"].head(30))
n = df[df["new cycle"] == True].shape[0]
print("Cycles identified: " + str(n))


# In[215]:


def analysis(df1, df2):
    A = df1.iloc[0, 1:] / df2.iloc[0, 1:] * 1000
    B = df1.iloc[1, 1:] / df2.iloc[1, 1:] * 1000
    C = df1.iloc[2, 1:7] / df2.iloc[2, 1:7] * 1000
    D = df1.iloc[3, 1:7] / df2.iloc[3, 1:7] * 1000
    
    avg1 = (B[0] + B[1]) / 2
    avg2 = (B[2] + B[3]) / 2
    avg3 = (B[4] + B[5]) / 2
    avg4 = (B[6] + B[7]) / 2
    avg5 = (B[8] + B[9]) / 2
    avg6 = (B[10] + B[11]) / 2
    avg7 = (D[0] + D[1]) / 2
    avg8 = (D[2] + D[3]) / 2
    avg9 = (D[4] + D[5]) / 2
    
    out = []
    
    out.append(A[0] - avg1)
    out.append(A[1] - avg1)
    out.append(A[2] - avg2)
    out.append(A[3] - avg2)
    out.append(A[4] - avg3)
    out.append(A[5] - avg3)
    out.append(A[6] - avg4)
    out.append(A[7] - avg4)
    out.append(A[8] - avg5)
    out.append(A[9] - avg5)
    out.append(A[10] - avg6)
    out.append(A[11] - avg6)
    out.append(C[0] - avg7)
    out.append(C[1] - avg7)
    out.append(C[2] - avg8)
    out.append(C[3] - avg8)
    out.append(C[4] - avg9)
    out.append(C[5] - avg9)
    
#     print(out)
    return out


# In[206]:


def findTime(s):
    time = s[s.find("(") + 1 : s.find(")")]
    minut = int(time[0 : time.find("min")])
    s = time.find("s")
    sec = 0
    if s != -1:
        sec = int(time[s - 4:s - 1])
    return minut * 60 + sec


# In[219]:


lines = []
times = []

for i, row in df.iterrows():
    if row["new cycle"] == True:
                
        print("Cycle : " + str(i))
        
        time = findTime(row[0])
        
        times.append(time)
            
        df1 = df.iloc[i + 2 : i + 7, 0 : 13].astype(int, errors = "ignore")
        df1.columns =  df1.iloc[0, :].to_list()
        df1.reset_index()
        df1.drop(df1.index[0], inplace = True)
                
        df2 = df.iloc[i + 13 : i + 19, 0 : 13].astype(int, errors = "ignore")
        df2.columns = df2.iloc[0, :].to_list()
        df2.drop(df2.index[0], inplace = True)

        vals = analysis(df1, df2)        
        lines.append(vals)
        
        print("Time : " + str(time) + " s")
        print(vals)
                


# In[225]:


out = pd.DataFrame(lines)
out["time s"] = times
cols = out.columns.tolist()
cols = cols[-1:] + cols[:-1]
out = out[cols]

print(out.head())


# In[226]:


out.to_excel("output.xlsx")
print("decoder ran successfully")

