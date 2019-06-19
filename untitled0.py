# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 18:32:06 2019

@author: praveen-ch
"""


def isBinary(c):
    f=False
    for i in c:
        if(i=='0' or  i=='1'):
            f= True
        else:
            f= False
    return f
c=input();
print(isBinary(c))

        
    
        