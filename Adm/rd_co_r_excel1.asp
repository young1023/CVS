<%
' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=Redemp_Raw_Rpt_"&Month(Now())&Year(now())&".xls"

%>

<!--#include file="include/SQLConn.inc" -->

<%


'**************
'Argument handler
'**************

	
From_Date      = Request.Form("From_Date")
To_Date        = Request.Form("To_Date")
Coupon_Type    = Request.Form("Coupon_Type")
Station        = Request.Form("Station")
Coupon_Batch   = Request.Form("Coupon_Batch")
Coupon_Number  = Request.Form("Coupon_Number")
Print_Excel    = Request.Form("Print_Excel")
Face_Value     = Request.Form("Face_Value")
Print_Excel    = Request.Form("Print_Excel")

if Print_Excel = "" then
   Print_Excel = "All"
end if

Excel_Type     = Request.Form("Excel_Type")


'**************

	

' Create a server recordset object

      set rs = server.createobject("adodb.recordset")
      response.write  ("Exec RedemptionReport2 '"&From_Date&"', '"&To_Date&"', '"&Station&"' ,'"&Coupon_Type&"', '"&Coupon_Batch&"', '"&Face_Value&"', '"&Coupon_Number&"','"&Excel_Type&"',  '"&Print_Excel&"' ")
      rs.open ("Exec RedemptionReport2 '"&From_Date&"', '"&To_Date&"', '"&Station&"' ,'"&Coupon_Type&"', '"&Coupon_Batch&"', '"&Face_Value&"', '"&Coupon_Number&"','"&Excel_Type&"',  '"&Print_Excel&"' ") ,  conn,3,1
        
%>

<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<body>
<Head>
<STYLE TYPE="text/css">
<!--

TD 
{
  color: black;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}

TD.caption
{
  color: red;
  font-family: verdana, Garamond, Times, sans-serif;
  FONT-SIZE: 9px;
  TEXT-ALIGN: left 
}
-->
</STYLE>
</head>

<div align="center">

<table BORDER="1" width="98%">

<tr bgcolor="#DFDFDF">

<td height="28">Present Date</td>
<td height="28">Present Time</td>

<td height="28">Station</td>
<td height="28">Coupon<br/>Type</td>
<td  height="28">Batch</td>
<td  height="28">Coupon Number</td>
<td  height="28">Product<br/>Type</td>
<td  height="28">Sale<br/>Amount</td>
<td>Issue Date</td>
<td >Expiry Date</td>
<td>Excel<br/> Type</td>

<td>Print Excel</td>
<td>Print Excel Date</td>
</tr>

<%
' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		
%>

<tr>

<td align=center width="95" height="28"><% = FormatDateTime(rs("Present Date"),2)%></td>
<td align=center width="95" height="28"><% = FormatDateTime(rs("Present Date"),4)%></td>

<td  height="28">
<% = rs("Station") %>
</td>
<td  height="28"><% = rs("Coupon Type") %>
</td>

<td height="28"><% = rs("Batch") %>
</td>

<td  height="28"><% = rs("Coupon number") %>
</td>

<td  height="28">
<% = rs("Product Type") %>
</td>

<td>
<% = rs("Sale Amount") %> 
</td>

<td >
<% = rs("Issue Date") %>
</td>

<td >
<% = rs("Expiry Date") %>
</td>


<td >
<% = rs("Excel Type") %>
</td>



<td >
<% = rs("Status") %>
</td>

<td >
<% = rs("Creation_Date") %>
</td>
</tr>

<%
' Move to the next record
rs.movenext
' Loop back to the do statement
loop %>
</table>

</div>

</body>
</html>

<%
' Close and set the recordset to nothing
rs.close
set rs=nothing
%>