f=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week20.xlsx',1,'K2:K113');
ic=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week20.xlsx',1,'O2:O113');
A=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week20.xlsx',1,'P2:S113')';
b=[6000;6000;6000;6000];
Aeq=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week20.xlsx',1,'U2:EB29');
beq=[1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1]; 
lb=zeros(112,1);
ub=ones(112,1);
[x,fval]=intlinprog(f,ic,A,b,Aeq,beq,lb,ub);