<%
' Tells the browser to open csv
Response.ContentType = "text/csv"
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Content-Disposition", "attachment; filename=Redempt_Raw_"&Month(Now())&"_"&Year(now())&".csv" 
%>
<!--#include file="include/SQLConn.inc" -->

<%


'**************
'Argument handler
'**************

	

From_Date      = Request.Form("From_Date")
From_Date      = "2015-01-01"
To_Date        = Request.Form("To_Date")
To_Date        = "2015-01-05"
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
      'response.write  ("Exec RedemptionReport2 '"&From_Date&"', '"&To_Date&"', '"&Station&"' ,'"&Coupon_Type&"', '"&Coupon_Batch&"', '"&Face_Value&"', '"&Coupon_Number&"','"&Excel_Type&"',  '"&Print_Excel&"' ")
      rs.open ("Exec RedemptionReport2 '"&From_Date&"', '"&To_Date&"', '"&Station&"' ,'"&Coupon_Type&"', '"&Coupon_Batch&"', '"&Face_Value&"', '"&Coupon_Number&"','"&Excel_Type&"',  '"&Print_Excel&"' ") ,  conn,3,1


Response.Write "Present Date, Present Time, Station, Coupon Type, Batch, Coupon Number, Product Type, Sale Amount, Issue Date, Expiry Date, Excel Type, Print Excel, Print Excel Date" & vbcrlf


' Move to the first record
rs.movefirst

' Start a loop that will end with the last record
do while not rs.eof
 
		

Response.Write """" &  FormatDateTime(rs("Present Date"),2) & ""","  

Response.Write """" &  FormatDateTime(rs("Present Date"),4) & ""","  

Response.Write """" &  rs("Station") & """," 

Response.Write """" &  rs("Coupon Type") & ""","  

Response.Write """" &  rs("Batch") & """," 

Response.Write """" & rs("Coupon number") & """," 

Response.Write """" & rs("Product Type") & """," 

Response.Write """" &  rs("Sale Amount") & ""","   

Response.Write """" &  rs("Issue Date") & """," 

Response.Write """" & rs("Expiry Date") & """," 

Response.Write """" &  rs("Excel Type") & """," 

Response.Write """" &  rs("Status") & """," 

Response.Write """" &  rs("Creation_Date") & """" & vbCrLf  

' Move to the next record
rs.movenext

' Loop back to the do statement
loop 



' Close and set the recordset to nothing
rs.close
set rs=nothing
%>