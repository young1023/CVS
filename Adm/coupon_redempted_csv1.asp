<%
' Tells the browser to open csv
Response.ContentType = "text/csv"
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Content-Disposition", "attachment; filename=Redempt_Rate_"&Month(Now())&"_"&Year(now())&".csv" 
%>
<!--#include file="include/SQLConn.inc" -->
<%

Coupon_Type    = Request.Form("Coupon_Type")
Coupon_Batch   = Request.Form("Coupon_Batch")
Coupon_Number  = Request.Form("Coupon_Number")
Face_Value     = Request.Form("Face_Value")
Range = trim(request.form("mid"))

response.write delid
      
   fsql = "Select * from Coupon_Redempted Where 1 = 1 "

   ' Coupon Type
   if Coupon_Type <> "" then

   fsql = fsql & " and Product_Type = '" & Coupon_Type & "' "
   
   end if

   ' Batch
   if Coupon_Batch <> "" then

   fsql = fsql & " and Batch = '" & Coupon_Batch & "' "
   
   end if

   ' Face Value
   if Face_Value <> "" then

   fsql = fsql & " and FaceValue like '%" & Face_Value & "%' "
   
   end if

  ' Check Coupon Number
   if Coupon_Number <> "" then

   fsql = fsql & " and Cast(Start_Range as float) <= Cast(" & Coupon_Number & " as float) "

   fsql = fsql & " and Cast(End_Range as float) >= Cast(" & Coupon_Number & " as float) "

   end if

   If Range <> "" Then

   fsql = fsql & "and RangeID in (" & Range & ")"

   End If


   fsql = fsql & " order by RangeID desc"

        'response.write fsql
        'response.end

        set frs = Conn.Execute(fsql)
	


Response.Write "Face Value ,Type, Batch, Start Range, End Range,  Total Used Coupons, Total Coupons Issued, Redemption Rate" & vbcrlf

 if not frs.EoF then
 
   frs.MoveFirst

   Do While not frs.EOF

   Response.Write """" &  frs("FaceValue") & """," 

   Response.Write """" &  frs("Product_Type") & """," 

   Response.Write """" & frs("Batch") & """," 
   Response.Write """" &  frs("Start_Range") & """," 
   Response.Write """" &  frs("End_Range") & """," 
  

      Set rs = server.createobject("adodb.recordset")

      'response.write  ("Exec Retrieve_Redemption_Rate2 '"&frs("Product_Type")&"', '"&frs("Batch")&"', '"&frs("FaceValue")&"' , '"&frs("Start_Range")&"', '"&frs("End_Range")&"'") 

	  rs.open ("Exec Retrieve_Redemption_Rate2 '"&frs("Product_Type")&"', '"&frs("Batch")&"', '"&frs("FaceValue")&"', '"&frs("Start_Range")&"', '"&frs("End_Range")&"'") ,  conn,3,1


Response.Write """" &  rs("RedemptionNo") & """," 
Response.Write """" &  rs("TotalNo") & """," 
Response.Write """" &  FormatNumber(rs("Rate"),1) & "%" & """" & vbCrLf  


   
   frs.movenext
  loop
 end if
 
frs.close
set frs=nothing
conn.close
set conn=nothing
%>



