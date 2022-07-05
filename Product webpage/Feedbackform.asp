
<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" /> 
    <title>Feedbackform</title>

    <h3>Thanks for the feedback, <%Response.Write(Request.Form("name")) %> .
        <br />
        Your comments will be kindly considered!</h3>
</head>
    
    <body>
    <%
        var gender = Request.Form("gender");
        var age = Request.Form("age");
        var drinks = Request.Form("drinks");
        var comment = Request.Form("comment");
        var rates = Request.Form("rates");
        var email = Request.Form("email");

var conn = new ActiveXObject("ADODB.Connection");
        conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/ISE2001project/Feedbackform.mdb'");
        var SqlString = "INSERT INTO Feedbackform(Gender, Age, Drinks, Comment, Rates, Email) VALUES('" + gender + "', '" + age + "', '" + drinks + "','" + comment + "', '" + rates + "', '" + email + "')";
conn.Execute(SqlString);
conn.Close;
%>
     </body>
</html>