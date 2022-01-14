import xlrd
from numpy import transpose
import math
import numpy as np

def atleast():
    xl = xlrd.open_workbook(r'C:\Users\LENOVO\Desktop\C\50.xlsx')
    table=xl.sheet_by_name("Sheet15")
    col=table.col_values(242,0,50)
    sum=0;point=0
    for i in range(0,50):
        sum=sum+col[i]
        if sum>=28200:
            point=i
            break
    return (point+1)


def eco():
    xl = xlrd.open_workbook(r'C:\Users\LENOVO\Desktop\C\price.xlsx')
    table = xl.sheet_by_name("Sheet2")
    cols=[];SUM=[]
    for i in range(0,129):
        col=table.col_values(i,0,29)
        sum=0
        for j in range(0,29):
            sum=sum+col[j]
        SUM.append(sum)
    SUM.sort()
    return SUM[0]

print(eco())
