<!-- #include file="includes/config.asp" -->
<!-- #include file="includes/header.asp" -->
<h2>Add New Task</h2>
<% If Request.Form("submit") <> "" Then
    title = Request.Form("title")
    description = Request.Form("description")
    dueDate = Request.Form("dueDate")
    status = Request.Form("status")
    
    If title <> "" And dueDate <> "" Then
        Set conn = GetConnection()
        sql = "INSERT INTO Tasks (Title, Description, DueDate, Status) VALUES ('" & _
              Replace(title, "'", "''") & "', '" & _
              Replace(description, "'", "''") & "', #" & _
              dueDate & "#, '" & _
              Replace(status, "'", "''") & "')"
        conn.Execute(sql)
        conn.Close
        Response.Redirect "index.asp"
    Else
        Response.Write "<p style='color:red'>Title and Due Date are required.</p>"
    End If
End If %>
<form method="post">
    <div class="form-group">
        <label>Title:</label>
        <input type="text" name="title">
    </div>
    <div class="form-group">
        <label>Description:</label>
        <textarea name="description"></textarea>
    </div>
    <div class="form-group">
        <label>Due Date:</label>
        <input type="text" name="dueDate" placeholder="MM/DD/YYYY">
    </div>
    <div class="form-group">
        <label>Status:</label>
        <select name="status">
            <option value="Pending">Pending</option>
            <option value="Completed">Completed</option>
        </select>
    </div>
    <input type="submit" name="submit" value="Add Task">
</form>
<!-- #include file="includes/footer.asp" -->
