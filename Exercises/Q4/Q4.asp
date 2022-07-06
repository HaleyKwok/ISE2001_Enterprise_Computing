<%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Q4</title>
    </head>
    <body>
 
        <br />
         <%
             // open the database
             var conn = new ActiveXObject("ADODB.Connection");
             var rs = new ActiveXObject("ADODB.Recordset");
             conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/Q4/Q4.mdb'");
             rs.Open("SELECT * FROM Products", conn);

             // get the input Product_ID from html
             var proid = Request.Form("Product_ID");
             var pid = String(proid);

             while (rs.EOF != true) {
                 var proid = rs("ProductID");
                 var ppid = String(proid);
                 if (pid == ppid) {
                     // the inputted pid is "", then display the corresponding picture in a certain size
                     if (pid == "1") { Response.write('<img src="1.jpg " style="max-height:200px; max-width:70%"/>') };
                     if (pid == "2") { Response.write('<img src="2.jpg " style="max-height:200px; max-width:70%"/>') };
                     if (pid == "3") { Response.write('<img src="3.jpg " style="max-height:200px; max-width:50%"/>') };
                     if (pid == "4") { Response.write('<img src="4.jpg " style="max-height:200px; max-width:50%"/>') };
                     if (pid == "5") { Response.write('<img src="5.jpg " style="max-height:200px; max-width:70%"/>') };
                     if (pid == "6") { Response.write('<img src="6.jpg " style="max-height:200px; max-width:70%"/>') };

                     // display the information extracted from database
                     Response.Write("<br>")
                     Response.Write("Product ID:")
                     Response.Write(rs("ProductID") + "<br>")
                     Response.Write("Product Name:")
                     Response.Write(rs("ProductName") + "<br>")
                     Response.Write("Product Descriptions:")
                     Response.Write(rs("ProductDescription") + "<br>")
                 }
                 // move to next record
                 rs.MoveNext();
             }
             rs.close();
             conn.Close();
%>
    </body>
</html>