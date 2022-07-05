<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" /> 
 
<title>VIP sign in</title>
</head>
<body>

  
<% 
 
    var conn = new ActiveXObject("ADODB.Connection");
    var rs = new ActiveXObject("ADODB.Recordset");
    conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/ISE2001project/VIPsignup.mdb'");

    var usern = Request.Form("Username");
    var passw = Request.Form("Pwd");


    var stringuser = String(usern);
    var stringpassw = String(passw);

    var found = false; 

    rs.Open("SELECT * FROM VIPs", conn);
    while (rs.EOF != true) {
        var us = rs("Username"); 
        var psw = rs("Pwd");   


        var userna = String(us);
        var passwor = String(psw); 

        if (userna == stringuser && passwor == stringpassw) { 
            found = true; 
            Response.Write("Hello! <BR>");
            Response.Write(rs("Firstname") + " " + rs("Lastname") + "<BR>");
            Response.Write("Your discount Code :");
            Response.Write(rs("Discount"));
            break;
        }

        rs.MoveNext();
    }

    if (found == false) {
        Response.Write('<script> alert("incorrect username or password")</script>'); 
    }

    rs.close();
    conn.Close();
    %>


</body> 
</html> 