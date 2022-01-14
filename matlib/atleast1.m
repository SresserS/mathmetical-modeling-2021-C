f=xlsread('C:\Users\LENOVO\Desktop\cnew\45.xlsx',1,'F2:F51');
ic=xlsread('C:\Users\LENOVO\Desktop\cnew\45.xlsx',1,'G2:G51');
A=-xlsread('C:\Users\LENOVO\Desktop\cnew\45.xlsx',1,'F2:F51')';
b=-28200;
Aeq=[1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1];
beq=28;
lb=zeros(50,1);
ub=ones(50,1);
c=[];temp=[];
for i=1:1000
    b=b-1;
    x=intlinprog([],ic,A,b,Aeq,beq,lb,ub);
    tf=isequal(x,temp);
    temp=x;
    if i>0&&~tf
        c=cat(2,c,x);
    end
end
xlswrite('C:\Users\LENOVO\Desktop\cnew\atleast.xlsx',c)