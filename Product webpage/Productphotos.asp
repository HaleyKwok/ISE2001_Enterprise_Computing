 <%@ Language="JavaScript" %> 
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    <style>
        table, th, td {
        border: 1px solid black;
        text-align: center;}
    </style>
</head>
<body>

    <table style="width:50%">
        <tr><td>
        
          <%
              var conn = new ActiveXObject("ADODB.Connection");
              var rs = new ActiveXObject("ADODB.Recordset");
              conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='P:/ISE2001project/productdescription.mdb'");
              rs.Open("SELECT * FROM Products", conn);

              var proid = Request.Form("P_ID");
              var pid = String(proid);

              while (rs.EOF != true) {
                  var proid = rs("P_ID");
                  var ppid = String(proid);
                  if (pid == ppid) {

                      if (pid == "1") { Response.write('<img src="1.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "2") { Response.write('<img src="2.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "3") { Response.write('<img src="3.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "4") { Response.write('<img src="4.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "5") { Response.write('<img src="5.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "6") { Response.write('<img src="6.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "7") { Response.write('<img src="7.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "8") { Response.write('<img src="8.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "9") { Response.write('<img src="9.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "10") { Response.write('<img src="10.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "11") { Response.write('<img src="11.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "12") { Response.write('<img src="12.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "13") { Response.write('<img src="13.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "14") { Response.write('<img src="14.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "15") { Response.write('<img src="15.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "16") { Response.write('<img src="16.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "17") { Response.write('<img src="17.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "18") { Response.write('<img src="18.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "19") { Response.write('<img src="19.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "20") { Response.write('<img src="20.jpg " style="max-height:200px; max-width:70%"/>') };
                      if (pid == "21") { Response.write('<img src="21.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "22") { Response.write('<img src="22.jpg " style="max-height:200px; max-width:50%"/>') };
                      if (pid == "23") { Response.write('<img src="23.jpg " style="max-height:200px; max-width:70%"/>') };
   

                      Response.Write("<br>")
                      Response.Write("Product ID:")
                      Response.Write(rs("P_ID") + "<br>")
                      Response.Write("Product Name:")
                      Response.Write(rs("ProductName") + "<br>")
                      Response.Write("Product Price:")
                      Response.Write(rs("ProductPrice") + "<br>")
                  }
                  rs.MoveNext();
              }
              rs.close();
              conn.Close();
%>

        

 </td ></tr >

** The picture above is for reference only. The real object should be considered as final.
  </table >

    </body>
</html>