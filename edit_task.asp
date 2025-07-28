<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/header.asp" -->
<h2>Edit Task</h2>
<%
id = Request.QueryString("id")
Set conn = GetConnection()
Set rs = conn.Execute("SELECT * FROM Tasks WHERE TaskID = " & id)
If Not rs.EOF Then
    title = rs("Title")
    description = rs("Description")
    dueDate = rs("DueDate")
    status = rs("Status")
End If
If Request.Form("submit") <> "" Then
    title = Request.Form("title")
    description = Request.Form("description")
    dueDate = Request.Form("dueDate")
    status = Request.Form("status")
    
    If title <> "" And dueDate <> "" Then
        sql = "UPDATE Tasks SET Title = '" & Replace(title, "'", "''") & "', " & _
              "Description = '" & Replace(description, "'", "''") & "', " & _
              "DueDate = #" & dueDate & "#, " & _
              "Status = '" & Replace(status, "'", "''") & "' " & _
              "WHERE TaskID = " & id
        conn.Execute(sql)
        conn.Close
        Response.Redirect "index.asp"
    Else
        Response.Write "<p style='color:red'>Title and Due Date are required.</p>"
    End If
End If
rs.Close
conn.Close
%>
<form method="post">
    <div class="form-group">
        <label>Title:</label>
        <input type="text" name="title" value="<%= title %>">
    </div>
    <div class="form-group">
        <label>Description:</label>
        <textarea name="description"><%= description %></textarea>
    </div>
    <div class="form-group">
        <label>Due Date:</label>
        <input type="text" name="dueDate" value="<%= dueDate %>">
    </div>
    <div class="form-group">
        <label>Status:</label>
        <select name="status">
            <option value="Pending" <%= If status = "Pending" Then "selected" Else "" %>>Pending</option>
            <option value="Completed" <%= If status = "Completed" Then "selected" Else "" %>>Completed</option>
        </select>
    </div>
    <input type="submit" name="submit" value="Update Task">
</form>
<!-- #include file="includes/footer.asp" -->
