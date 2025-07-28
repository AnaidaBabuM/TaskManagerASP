<%
Const DB_PATH = "db/tasks.mdb"
Function GetConnection()
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DB_PATH)
    Set GetConnection = conn
End Function
%>
