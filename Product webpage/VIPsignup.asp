<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" /> 
    <title>VIP sign up</title>

    <h3>Thanks for signing up, <%Response.Write(Request.Form("uname")) %> .
        <br />
        Your discount code will be available soon.</h3>
</head>
    
    <body>
    <%
        var uname = Request.Form("uname");
        var pwd = Request.Form("pwd");
        var lname = Request.Form("lname");
        var fname = Request.Form("fname");
        var number = Request.Form("number");

        var conn = new ActiveXObject("ADODB.Connection");
        conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/ISE2001project/VIPsignup.mdb'");
        var SqlString = "INSERT INTO VIPs(Username, Pwd, Lastname, Firstname, Phone) VALUES('" + uname + "', '" + pwd + "', '" + lname + "', '" + fname + "', '" + number + "')";

        conn.Execute(SqlString);
conn.Close;
%>
     </body>
</html>