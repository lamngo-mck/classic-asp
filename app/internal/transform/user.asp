<%
Function ConvertUserJson(user)
	Set oJSON = New aspJSON

	With oJSON.data

		.Add "id", user.Id                     
		.Add "username", user.Username
		.Add "email", user.Email

	End With

	set ConvertUserJson = oJSON.JSONoutput()   
End Function
%>