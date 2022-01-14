import xlrd
from numpy import transpose
import math
import numpy as np

xl = xlrd.open_workbook(r'C:\Users\LENOVO\Desktop\C\ABC.xlsx')
table1=xl.sheets()[0]
table2=xl.sheets()[1]
i=0;j=0;times=0;t=0
A_rowsD=[];A_rowsG=[]
A_Q_rate=[];A_W_rate=[];A_G_rate=[];A_F=[]
B_rowsD=[];B_rowsG=[]
B_Q_rate=[];B_W_rate=[];B_G_rate=[];B_F=[]
C_rowsD=[];C_rowsG=[]
C_Q_rate=[];C_W_rate=[];C_G_rate=[];C_F=[]

def compute(begin,end):
#求缺货率和完成率
    for i in range(begin,end): #1,147
    #缺货率
        t=0;times=0
        A_rowD=table1.row_values(i,2,242)
        A_rowG=table2.row_values(i,2,242)
        A_rowsD.append(A_rowD)
        A_rowsG.append(A_rowG)
        for j in range(0,240):
            if A_rowG[j]<A_rowD[j]:
                times=times+1
            if A_rowD[j]>0:
                t=t+1
        A_Q_rate.append(times/t)

    #完成率
        sumrate = 0;t = 0
        for j in range(0, 240):
            if A_rowD[j] > 0:
                rate = A_rowG[j] / A_rowD[j]
                t = t + 1
                sumrate = sumrate + rate
        sumrate = sumrate / t
        A_W_rate.append(sumrate)

    #方差
        npa=np.array(A_rowG)
        var=np.var(npa)
        A_F.append(var)


#求供货率
    A_colD=[];A_colG=[];weeks=[]
    for j in range(2,242):
        sum=0;week=[]
        A_colD=table1.col_values(j,begin,end)
        A_colG=table2.col_values(j,begin,end) #1,147
        for i in range(0,end-begin): #0,146
            sum=sum+A_colG[i]
        for i in range(0,end-begin):
            rate=A_colG[i]/sum
            week.append(rate)
        weeks.append(week)

    transposed = transpose(weeks).tolist()
    print(len(weeks[0]))
    for i in range(0,end-begin): #0,146
        sum=0
        for j in range(0,240):
            sum=sum+transposed[i][j]
        A_G_rate.append(sum/240)

    return A_Q_rate,A_W_rate,A_G_rate,A_F

C_Q_rate,C_W_rate,C_G_rate,C_F=compute(281,403)
A_Q_rate=[];A_W_rate=[];A_G_rate=[];A_F=[]
B_Q_rate,B_W_rate,B_G_rate,B_F=compute(147,281)
A_Q_rate=[];A_W_rate=[];A_G_rate=[];B_F=[]
A_Q_rate,A_W_rate,A_G_rate,A_F=compute(1,147)

A_Q_rate.extend(B_Q_rate)
A_Q_rate.extend(C_Q_rate)
A_W_rate.extend(B_W_rate)
A_W_rate.extend(C_W_rate)
A_G_rate.extend(B_G_rate)
A_G_rate.extend(C_G_rate)
A_F.extend(B_F)
A_F.extend(C_F)

arr=[]
arr.append(A_Q_rate)
arr.append(A_W_rate)
arr.append(A_G_rate)
arr.append(A_F)
def huise(arr): #arr是矩阵ABC合，四项指标
    #缺货，极小指标
    arr1=[];min_=[];max_=[];arr1s=[]
    for i in range(0,4):
        min_.append(min(arr[i])) #行是指标
        max_.append(max(arr[i]))
    arr0=[min_[0],max_[1],max_[2],min_[3]]
    #无量纲矩阵
    for i in range(0,4):
        for j in range(0,402):
            if i==0:
                x=(arr[i][j]-min_[i])/(max_[i]-min_[i])
                arr1.append(x)
            else:
                x = (max_[i]-arr[i][j]) / (max_[i] - min_[i])
                arr1.append(x)
        arr1s.append(arr1)
        arr1=[]


    #关联系数矩阵
    gl=[];gls=[]
    temp=[]
    for i in range(0,402):
        nparr1s=np.array(arr1s).transpose()
        nparr0=np.array(arr0)
        temp.append(min(abs(nparr0-nparr1s[i])))
    MIN=min(temp)
    temp=[]
    for i in range(0,4):
        nparr1s = np.array(arr1s).transpose()
        nparr0 = np.array(arr0)
        temp.append(max(abs(nparr0 - nparr1s[i])))
        '''temp.append(max(list(map(lambda x: x[0]-x[1], zip(arr1s[i],arr0)))))'''
    MAX=max(temp)
    q=0.5
    for i in range(0,4):
        for j in range(0,402):
            xs=(MIN+q*MAX)/(abs(arr0[i]-arr1s[i][j])+q*MAX) #系数
            gl.append(xs)
        gls.append(gl)

    #权重
    p=[];ps=[]
    for i in range(0,4):
        sum_ = sum(arr1s[i])
        for j in range(0,402):
            p.append(arr1s[i][j]/sum_)
        ps.append(p)
    #熵
    E=[];k=1/math.log(4)
    for i in range(0,4):
        E.append(0)
        for j in range(0,402):
            if ps[i][j]>0:
                E[i]=E[i]-k*ps[i][j]*math.log(ps[i][j])
    d=[];sum_=0;w=[]
    for i in range(0,4):
        d.append(1-E[i])
        sum_=sum_+d[i]
    for i in range(0,4):
        w.append(d[i]/sum_)

    #评估
    R=[]
    for j in range(0,402):
        a=0
        for i in range(0,4):
            a=a+w[i]*gls[i][j]
        R.append(a)
    return R

r=[]
r=huise(arr)
print(r)





