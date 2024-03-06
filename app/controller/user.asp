
<!-- #include virtual="app\service\user.asp" -->

<%

id = request.QueryString("id")
Set data = UserService.GetUserById(id)

response.Write data
response.ContentType "application/json"

%>