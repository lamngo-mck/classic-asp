<!-- #include virtual="\app\repository\user.asp" -->
<!-- #include virtual="\app\internal\transform\user.asp" -->

<%
Class UserService_Class
	Public Function GetById(id)
		user = UserRepository.FindById(id)
		data = ConvertUserJson(user)

		set GetUserById = data
	End Function
End Class

dim UserService__Singleton
Function UserService()
    If IsEmpty(UserService__Singleton) then
        set UserService__Singleton = new UserService_Class
    End If
    set UserService = UserService__Singleton
End Function
%>