<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/header.asp" -->
<h2>Task List</h2>
<%
Set conn = GetConnection()
Set rs = conn.Execute("SELECT TaskID, Title, DueDate, Status FROM Tasks ORDER BY DueDate")
%>
<table>
    <tr>
        <th>Title</th>
        <th>Due Date</th>
        <th>Status</th>
        <th>Actions</th>
    </tr>
<%
Do While Not rs.EOF
%>
    <tr>
        <td><a href="view_task.asp?id=<%= rs("TaskID") %>"><%= rs("Title") %></a></td>
        <td><%= rs("DueDate") %></td>
        <td><%= rs("Status") %></td>
        <td>
            <a href="edit_task.asp?id=<%= rs("TaskID") %>">Edit</a> |
            <a href="delete_task.asp?id=<%= rs("TaskID") %>" onclick="return confirm('Are you sure?')">Delete</a>
        </td>
    </tr>
<%
    rs.MoveNext
Loop
rs.Close
conn.Close
%>
</table>
<!-- #include file="includes/footer.asp" -->
