f=xlsread('C:\Users\LENOVO\Desktop\cnew\45.xlsx',1,'E2:E51');
ic=xlsread('C:\Users\LENOVO\Desktop\cnew\45.xlsx',1,'G2:G51');
A=-xlsread('C:\Users\LENOVO\Desktop\cnew\45.xlsx',1,'F2:F51')';
b=-28200;
Aeq=[1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1];
beq=[];
lb=zeros(50,1);
ub=ones(50,1);
[x,fval]=intlinprog(f,ic,A,b,[],[],lb,ub);