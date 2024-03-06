
<!-- #include virtual="app\service\user.asp" -->

<%
Class UserController

    Public Sub GetById
        id = request.QueryString("id")
        Set data = UserService.GetById(id)

        response.Write data
        response.ContentType "application/json"
    End Sub

End Class
%>