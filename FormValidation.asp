<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>ASP Demo</title>
        <%
            set conn=Server.CreateObject("ADODB.Connection")
            conn.Provider="Microsoft.SQLSERVER.CE.OLEDB.4.0"
            conn.Open "C:\Users\Aaron\Documents\My Web Sites\EmptySite\App_Data\Northwind.sdf"
            set rs = Server.CreateObject("ADODB.recordset")
            rs.Open "SELECT * FROM person", conn
        %>
    </head>

    <body>

        <form method="get" action="Index.asp">
            <input type="text" name="fname" placeholder="First Name"><br>
            <input type="text" name="lname" placeholder="Last Name"><br>
            <input type="submit" value="Submit"><br>
        </form>
        <p>
            <% 
                Dim FirstName
                Dim LastName
                FirstName = Request.QueryString("fname")
                LastName = Request.QueryString("lname")
                If FirstName <> "" AND LastName <> "" Then
                    Response.Write("Hello: " & FirstName & " " & LastName & "!!!")
                Else
                    Response.Write("Please fill out both the first and last name fields!!")
                End IF
            %>
        </p>

        <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.2/jquery.min.js"></script>
        <script type = "text/javascript"src = "https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.3/jquery-ui.min.js"></script>
        <script src="Script.js"></script>
    </body>
</html>
