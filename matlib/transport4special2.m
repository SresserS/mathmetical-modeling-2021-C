f=xlsread('C:\Users\LENOVO\Desktop\cnew\transport4special\week2.xlsx',1,'K2:K1605');
ic=xlsread('C:\Users\LENOVO\Desktop\cnew\transport4special\week2.xlsx',1,'O2:O1605');
A=xlsread('C:\Users\LENOVO\Desktop\cnew\transport4special\week2.xlsx',1,'Q2:T1605')';
b=[6000;6000;6000;6000];
Aeq=xlsread('C:\Users\LENOVO\Desktop\cnew\transport4special\week2.xlsx',1,'W2:BJN402');
beq=xlsread('C:\Users\LENOVO\Desktop\cnew\transport4special\week2.xlsx',1,'F2:F402');
lb=zeros(1604,1);
ub=ones(1604,1);
[x,fval]=intlinprog(f,ic,A,b,Aeq,beq,lb,ub);