<%
Class Database_Class
    Private m_connection
    Private m_connection_string
    
    '---------------------------------------------------------------------------------------------------------------------
    Public Sub Initialize(connection_string)
        m_connection_string = connection_string
    End Sub

    '---------------------------------------------------------------------------------------------------------------------
    Public Function Query(sql, params)
        dim cmd : set cmd = server.createobject("adodb.command")
        set cmd.ActiveConnection = Connection
        cmd.CommandText = sql
        
        dim rs
        
        If IsArray(params) then
            set rs = cmd.Execute(, params)
        ElseIf Not IsEmpty(params) then
            set rs = cmd.Execute(, Array(params))
        Else
            set rs = cmd.Execute()
        End If
        
        set Query = rs
    End Function  

    Private Function Connection
        if not isobject(m_connection) then 
            set m_connection = Server.CreateObject("adodb.connection")
            m_connection.open m_connection_string
        end if
        set Connection = m_connection
    End Function
End Class
%>