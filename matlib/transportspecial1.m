f=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week1.xlsx',1,'M2:M146');
ic=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week1.xlsx',1,'R2:R146');
A=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week1.xlsx',1,'T2:X146')';
b=[6000;6000;6000;6000;6000];
Aeq=xlsread('C:\Users\LENOVO\Desktop\cnew\transportspecial\Week1.xlsx',1,'Y2:FM30');
beq=[1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1]; 
lb=zeros(145,1);
ub=ones(145,1);
[x,fval]=intlinprog(f,ic,A,b,Aeq,beq,lb,ub);