<%@ Language="VBScript" %>

<%
    Response.expires = -1
    dim sql
    sql = "SELECT * FROM PERSON WHERE name="
   	sql = sql & "'" & request.querystring("q") & "'"

    set conn=Server.CreateObject("ADODB.Connection")
    conn.Provider="Microsoft.SQLSERVER.CE.OLEDB.4.0"
    conn.Open "C:\Users\Aaron\Documents\My Web Sites\EmptySite\App_Data\Northwind.sdf"
    set rs = Server.CreateObject("ADODB.recordset")
    rs.Open sql, conn

    Response.Write("<table>")
    do until rs.EOF
    	for each x in rs.Fields
	    	Response.Write("<tr><td><b>" & x.name & "</b></td>")
	    	Response.Write("<td>" & x.value & "</td></tr>")
	    next
	    rs.MoveNext
    loop
    Response.Write("</table>")
%>
