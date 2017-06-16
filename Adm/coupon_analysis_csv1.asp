<%
' Tells the browser to open excel
Response.ContentType = "text/csv"
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Content-Disposition", "attachment; filename=Raw_Rpt_"&Month(Now())&"_"&Year(now())&".csv" 
%>
<!--#include file="include/SQLConn.inc" -->
<%

Coupon_Type    = Request.Form("Coupon_Type")
Coupon_Batch   = Request.Form("Coupon_Batch")
Start_Range    = Request.Form("Start_Range")
End_Range      = Request.Form("End_Range")
Face_Value     = Request.Form("Face_Value")
Excel_Type     = Request.Form("Excel_Type")
Expiry_Date    = Request.Form("Expiry_Date")
Coupon_Number  = Request.Form("Coupon_Number")


      
    fsql = "select * from CouponRequest where 1 =1  "

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


  ' Check Excel type
   if Excel_type <> "" then

   fsql = fsql & " and Excel_type = '" & Excel_type & "' "
   
   end if


       fsql = fsql & " order by Issue_Date desc"

        'response.write fsql
        'response.end

        set frs = Conn.Execute(fsql)
	


Response.Write "Face Value ,Type, Batch, Start Range, End Range, Digital, Category, Expiry Date, Issue Date, Excel Type, Total Used Coupons, Total Coupons Issued, Redemption Rate" & vbcrlf

 if not frs.EoF then
 
   frs.MoveFirst

   Do While not frs.EOF

   Response.Write """" &  frs("FaceValue") & """," 

   Response.Write """" &  frs("Product_Type") & """," 

   Response.Write """" & frs("Batch") & """," 
Response.Write """" &  frs("Start_Range") & """," 
Response.Write """" &  frs("End_Range") & """," 
Response.Write """" &  frs("Digital") & """," 
Response.Write """" &  frs("Category") & """," 
Response.Write """" &  frs("Expiry_Date") & """," 
Response.Write """" &  frs("Issue_Date") & """," 
Response.Write """" &  frs("Completed") & """," 
Response.Write """" &  frs("Excel_Type") & """," 


      Set rs = server.createobject("adodb.recordset")

      'response.write  ("Exec Retrieve_Redemption_Rate '"&frs("Product_Type")&"', '"&frs("Batch")&"', '"&frs("FaceValue")&"' , '"&frs("Start_Range")&"', '"&frs("End_Range")&"', '"&frs("Excel_Type")&"'") 

	  rs.open ("Exec Retrieve_Redemption_Rate '"&frs("Product_Type")&"', '"&frs("Batch")&"', '"&frs("FaceValue")&"', '"&frs("Start_Range")&"', '"&frs("End_Range")&"', '"&frs("Excel_Type")&"'") ,  conn,3,1


Response.Write """" &  rs("RedemptionNo") & """," 
Response.Write """" &  rs("TotalNo") & """," 
Response.Write """" &  FormatNumber(rs("Rate"),1) & """" & vbCrLf  


   
   frs.movenext
  loop
 end if
 
frs.close
set frs=nothing
conn.close
set conn=nothing
%>



