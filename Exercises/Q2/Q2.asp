<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" /> 
    <title>Q2</title>
</head>
<body>
<h1>Covid-19 Vaccination</h1>

<p>Your booking has been confirmed!</p>
Name: <%Response.Write(Request.Form("firstname")) %>  <%Response.Write(Request.Form("lastname")) %>
<br />
Vaccine: <%Response.Write(Request.Form("vaccine")) %>
<br />
Reference number: <span id="reference"></span>
        <script>
            //the reference number range
            var characters = "0123456789abcdefghijklmnopqrstuvwxyz"
            //the length of reference number
            var length = 7

            function generate() {
                var ref = ""
                var n = 0
                var randomnumber = 0
                while (n < length) {
                    n++
                    randomnumber = Math.floor(characters.length * Math.random());
                    ref += characters.substring(randomnumber, randomnumber + 1)
                }
                return ref
            }
            //display the number
            document.getElementById("reference").innerHTML = generate();
        </script>
<br />
First dose: <%Response.Write(Request.Form("date")) %>

<p>Please remember to bring your identity document for vaccination.</p>



 <%
     // get the information from the html
     var firstname = Request.Form("firstname");
     var lastname = Request.Form("lastname");
     var vaccine = Request.Form("vaccine");
     var date = Request.Form("date");
     var age = Request.Form("age");
     var reference = Request.Form("reference");

     // open the database
     var conn = new ActiveXObject("ADODB.Connection");
     conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/Q2/Q2.mdb'");

     //insert the information correspondingly
     var SqlString = "INSERT INTO Bookings(FirstName, LastName, Age, Vaccine, Reference, FirstDose) VALUES('" + firstname + "', '" + lastname + "', '" + age + "','" + vaccine + "', '" + reference + "', '" + date + "')";
     conn.Execute(SqlString);
     conn.Close;
%>

</body>
</html>

