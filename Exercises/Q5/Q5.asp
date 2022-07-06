<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" /> 
    <title>Q5</title>
    <h1 style="color: #B22222;">PolyU CAR Subject Feedback</h1>
    <h3>Thanks for the feedback, <%Response.Write(Request.Form("name")) %> . 
        <br />
        Your comments will be kindly considered!</h3>
</head>
    
    <body>
    <%
      
        var name = Request.Form("name");
        var id = Request.Form("id");
        var gender = Request.Form("gender");
        var year = Request.Form("year");
        var A = Request.Form("A");
        var Arates = Request.Form("Arates");
        var B = Request.Form("B");
        var Brates = Request.Form("Brates");
        var C = Request.Form("C");
        var Crates = Request.Form("Crates");
        var D = Request.Form("D");
        var Drates = Request.Form("Drates");

// open database
var conn = new ActiveXObject("ADODB.Connection");
        conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/Q5/Q5.mdb'");
        // insert information into database
        var SqlString = "INSERT INTO Feedback(SName,SID, Gender,StudyYear, CARA, CARA_Rates, CARB, CARB_Rates, CARC, CARC_Rates, CARD, CARD_Rates) VALUES('" + name + "', '" + id + "', '" + gender + "','" + year + "', '" + A + "', '" + Arates + "', '" + B + "', '" + Brates + "', '" + C + "', '" + Crates + "', '" + D + "', '" + Drates + "')";

conn.Execute(SqlString);
conn.Close;
%>
     </body>

    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />


    </div>
</html>