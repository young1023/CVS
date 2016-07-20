<%
' Tells the browser to open excel

Response.ContentType = "application/vnd.ms-excel"
Response.addHeader "content-disposition","attachment;filename=Redemp_Rpt_"&Month(Now())&Year(now())&".xls"

%>

<!--#include file="include/SQLConn.inc" -->

<%


'**************
'Argument handler
'**************

	
From_Date      = Request.Form("From_Date")
To_Date        = Request.Form("To_Date")
Coupon_Type    = Request.Form("Coupon_Type")
Coupon_Batch   = Request.Form("Coupon_Batch")
Start_Range    = Request.Form("Start_Range")
End_Range      = Request.Form("End_Range")
Face_Value     = Request.Form("Face_Value")
	

' Create a server recordset object

      'set rs = server.createobject("adodb.recordset")
      'response.write  ("Exec RedemptionReport1 '"&From_Date&"', '"&To_Date&"' ") 
	  'rs.open ("Exec RedemptionReport1 '"&From_Date&"', '"&To_Date&"' ") ,  conn,3,1

      fsql = "SELECT  "

      fsql = fsql & " Convert(datetime, Present_Date,103) as [Present Date] , m.RequestedID as Station, "

      fsql = fsql & " m.Coupon_Number as [Coupon Number], "

      fsql = fsql & " Convert(varchar, c.Expiry_Date,103) as [Expiry Date] "

      fsql = fsql & " From MasterCoupon m Left Join CouponRequest c on m.Coupon_Type = c.Product_Type and "

      fsql = fsql & " m.Coupon_batch = c.Batch and Cast(m.Face_Value as float) = Cast(c.FaceValue as float) and "

      fsql = fsql & " c.Start_Range <= m.Coupon_Number and c.End_Range >= m.Coupon_Number where "

      fsql = fsql & " (Datediff(day, Present_Date, '"&From_Date&"') < = 0 or '"&From_Date&"' = '') "

      fsql = fsql & " and (Datediff(day, Present_Date, '"&To_Date&"') >= 0 or '"&To_Date&"' = '') "

      fsql = fsql & " and (m.RequestedID = '"&Station&"'  or '"&Station&"' = '') and "

      fsql = fsql & " (m.Coupon_Type = '"&Coupon_Type&"'  or '"&Coupon_Type&"' = '') and "

      fsql = fsql & " (m.Coupon_Batch = '"&Coupon_Batch&"' or '"&Coupon_Batch&"' = '') and "

      fsql = fsql & " (Cast(m.Face_Value as float)= '"&Face_Value&"' or '"&Face_Value&"' = '') and "

      fsql = fsql & " (m.Coupon_Number >= '"&Start_Range&"' or '"&Start_Range&"' = '') and "

      fsql = fsql & " (m.Coupon_Number <= '"&End_Range&"' or '"&End_Range&"' = '') "

      fsql = fsql & " group by Convert(datetime, Present_Date,103), m.requestedid, "

      fsql = fsql & " m.Coupon_Number, "

      fsql = fsql & " c.Expiry_Date "

      fsql = fsql & "  Order by Convert(datetime, Present_Date,103), m.Requestedid, "

      fsql = fsql & " m.Coupon_Number, c.Expiry_date"

      Set frs = Conn.Execute(fsql)

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
<tr>
<td height="28">Present Date</td>
<td height="28">Station</td>
<td height="28">Coupon Number</td>
<td height="28">Expiry Date</td>

</tr>

<%
' Move to the first record
frs.movefirst

' Start a loop that will end with the last record
do while not frs.eof
 
		
%>

   <tr>

<td align=center width="95" height="28"><% = frs("Present Date")%></td>
<td  height="28"><% = frs("Station") %>
</td>

<td height="28"><% = frs("Coupon Number") %>
</td>

<td height="28"><% = frs("Expiry Date") %>
</td>
</tr>

<%
' Move to the next record
frs.movenext
' Loop back to the do statement
loop %>
</table>

</div>

</body>
</html>

<%
' Close and set the recordset to nothing
frs.close
set frs=nothing
%>