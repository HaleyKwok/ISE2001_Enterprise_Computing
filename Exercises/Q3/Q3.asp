
<%@ Language="JavaScript" %> <!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" /> 
 
<title>Q3</title>
</head>
<body>

  
<% 
    // open the database
    var conn = new ActiveXObject("ADODB.Connection");
    var rs = new ActiveXObject("ADODB.Recordset");
    conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/Q3/Q3.mdb'");

    // get the value of username and password that input from the html
    var usern = Request.Form("Username");
    var passw = Request.Form("Pwd");

    // set the username and password requested into string
    var stringuser = String(usern);
    var stringpassw = String(passw);

    // check the password
    var found = false; 

    rs.Open("SELECT * FROM Users", conn);
    while (rs.EOF != true) {
        var us = rs("Username"); 
        var psw = rs("Pwd");   

        // set as string
        var userna = String(us);
        var passwor = String(psw); 

    //compare the username and password between input and database
        if (userna == stringuser && passwor == stringpassw) { 
            found = true; 
            Response.Write("Hello <BR>");
            Response.Write(rs("FirstName") + " " + rs("LastName") + "<BR>");
            break;       //stop the loop when it match the data
        }
        rs.MoveNext();   //move to next record
    } if (found == false) {
        Response.Write('<script> alert("incorrect username or password")</script>'); // alert the incorrect match
    }

    rs.close();
    conn.Close();
    %>


 
</body> 
</html> 