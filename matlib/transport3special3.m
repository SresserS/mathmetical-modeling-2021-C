f=xlsread('C:\Users\LENOVO\Desktop\cnew\transport3special\week3.xlsx',1,'I2:I185');
ic=xlsread('C:\Users\LENOVO\Desktop\cnew\transport3special\week3.xlsx',1,'L2:L185');
A=xlsread('C:\Users\LENOVO\Desktop\cnew\transport3special\week3.xlsx',1,'O2:Q185')';
b=[6000;6000;6000];
Aeq=xlsread('C:\Users\LENOVO\Desktop\cnew\transport3special\week3.xlsx',1,'S2:EZ47');
beq=[1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1;1]; 
lb=zeros(184,1);
ub=ones(184,1);
[x,fval]=intlinprog(f,ic,A,b,Aeq,beq,lb,ub);