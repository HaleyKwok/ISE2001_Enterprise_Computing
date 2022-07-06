<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">

<head>
<meta charset="utf-8" /> 
    <title>asp</title>
</head>

<body>
    <h1>Welcome!</h1>
    Your customer ID: <%Response.Write(Request.Form("CustomerID") + "<BR>") %>
    Your Name: <% Response.Write(Request.Form("FirstName")) %>
                <% Response.Write("" + Request.Form("LastName")) %>

<%
    var conn = new ActiveXObject("ADODB.Connection");
    var rs = new ActiveXObject("ADODB.Recordset");
    conn.open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/20060241d/20060241d.mdb'");
    rs.open("SELECT * FROM Customers", conn);
%>
    <table style="width:50%">
        <tr><td>
        <img id="demo" height="200">

            <script>
                document.getElementById("demo").src = <%=Request.Form("C_ID") %> + ".jpg"
            </script>
 </td ></tr >
  </table >


    <%
    rs.close(); conn.Close();
 %> 
    

</body>
</html>