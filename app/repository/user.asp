<!--#include file="../database/dal.asp"-->
<!--#include file="../utils/mapper.asp"-->

<%
Class UserRepository_Class
    Public Function FindById(id)
        dim sql : sql = "SELECT Id, Username, Email " &_
                        "FROM Users WHERE Id = ? "
        dim rs : set rs = DAL.Query(sql, id)

        dim order : set order = Automapper.AutoMap(rs, new UserModel_Class)

        set FindById = order
        Destroy rs
    End Function
End Class

dim UserRepository__Singleton
Function UserRepository()
    If IsEmpty(UserRepository__Singleton) then
        set UserRepository__Singleton = new UserRepository_Class
    End If
    set UserRepository = UserRepository__Singleton
End Function
%>