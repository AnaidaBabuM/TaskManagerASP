<!-- #include file="includes/config.asp" -->
<%
id = Request.QueryString("id")
If id <> "" Then
    Set conn = GetConnection()
    conn.Execute("DELETE FROM Tasks WHERE TaskID = " & id)
    conn.Close
End If
Response.Redirect "index.asp"
%>
