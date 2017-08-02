<%
' Tells the browser to open csv
Response.ContentType = "text/csv"
Response.AddHeader "Cache-Control", "no-cache"
Response.AddHeader "Content-Disposition", "attachment; filename=Redempt_Rate_"&Month(Now())&"_"&Year(now())&".csv" 
%>
<!--#include file="include/SQLConn.inc" -->
<%

id = request("id")

      
       fsql = "Select a.LeafLetNo as 'LeafLetNo' , a.FaceValue as 'FaceValue1', "

       fsql = fsql & "a.Coupon_Type as 'Coupon_Type1', a.Batch as 'Batch1', "

       fsql = fsql & "a.Start_Range as 'Start_Range1', a.End_Range as 'End_Range1', "

       fsql = fsql & " a.NoOfCoupon as 'TotalNo1', b.FaceValue as 'FaceValue2', "

       fsql = fsql & "b.Coupon_Type as 'Coupon_Type2', b.Batch as 'Batch2', "

       fsql = fsql & "b.Start_Range as 'Start_Range2', b.End_Range as 'End_Range2', "

       fsql = fsql & "b.NoOfCoupon as 'TotalNo2', TotalNo = (a.NoOfCoupon + b.NoOfCoupon ), "

       fsql = fsql & "TotalRedeem = (a.Redeem + b.Redeem) "

       fsql = fsql & "from Coupon_Redemption a inner join coupon_Redemption b "

       fsql = fsql & "on a.leafletid = b.leafletid and a.leafletno = b.leafletno "

       fsql = fsql & "and a.id = (b.id - 1) where a.leafletid =" & ID

       fsql = fsql & " order by a.LeafLetID "

   

        'response.write fsql
        'response.end

        set frs = Conn.Execute(fsql)
	


Response.Write "Face Value ,Type, Batch, Start Range, End Range, Face Value ,Type, Batch, Start Range, End Range, Total Used Coupons, Total Coupons Issued, Redemption Rate" & vbcrlf

 if not frs.EoF then
 
   frs.MoveFirst

   Do While not frs.EOF

   LeafLetNo = frs("LeafLetNo")

   Response.Write """" &  frs("FaceValue1") & """," 

   Response.Write """" &  frs("Coupon_Type1") & """," 

   Response.Write """" & frs("Batch1") & """," 

   Response.Write """" &  frs("Start_Range1") & """," 

   Response.Write """" &  frs("End_Range1") & """," 

   Response.Write """" &  frs("FaceValue2") & """," 

   Response.Write """" &  frs("Coupon_Type2") & """," 

   Response.Write """" & frs("Batch2") & """," 

   Response.Write """" &  frs("Start_Range2") & """," 

   Response.Write """" &  frs("End_Range2") & """," 

  Response.Write """" &  frs("TotalRedeem")  & """," 

  Response.Write """" &  frs("TotalNo") & """," 

  response.write FormatNumber(frs("TotalRedeem")/frs("TotalNo") * 100,1)  & "%" & """" & vbCrLf  

     
   frs.movenext
  loop
 end if
 
frs.close
set frs=nothing
conn.close
set conn=nothing
%>



