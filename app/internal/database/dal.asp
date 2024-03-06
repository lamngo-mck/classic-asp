<%
dim DAL__Singleton : set DAL__Singleton = Nothing

Function DAL()
    If DAL__Singleton is Nothing then
        set DAL__Singleton = new Database_Class
        DAL__Singleton.Initialize "Provider=SQLNCLI11;Data Source=EC2AMAZ-IV0IC7L\SQLEXPRESS;Initial Catalog=test;User ID=sa;Password=123456;Trust Server Certificate=True;"
    End If
    set DAL = DAL__Singleton
End Function
%>