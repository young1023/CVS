 <!--#include file="include/SQLConn.inc" -->

<%


   Set Rs1 = Server.CreateObject("Adodb.Recordset")
   Set Rs2 = Server.CreateObject("Adodb.Recordset")

   StationID = Session("StationID")

   UserID    = Session("UserID")

   SecLevel  = Session("SecLevel")

   Barcode    = Request("Barcode")

   ProductType = Request("ProductType")

   CarPlate   = Request("CarPlate")

   response.write Stationid

   response.write userid

   response.write seclevel

    Rs1.open ("Exec Checkrange '"&Barcode&"'") ,  conn,3,1

   'Response.write ("Exec Checkrange '"&Barcode&"'")

   'Response.end

   If Rs1.EoF Then

           Message = "Coupon is not in the Range."

 
   Else


       If Rs1("Expiry_Date") < DateValue(Now) Then

  
          Message = "Coupon expired"


       End if

   End if


          Rs1.close
          set Rs1 = nothing

          Rs2.open ("Exec CheckCouponExist '"&Barcode&"'") ,  conn,3,1

          Response.write ("Exec CheckCouponExist '"&Barcode&"'")

          iCount = Rs2("iCount")

          Rs2.close
          set Rs2 = nothing

         If iCount = 0 Then

         iMessage = "Success."

         Set Rs3 = Server.CreateObject("Adodb.Recordset")

         Response.Write ("Exec InsertCoupon '"&Barcode&"' , '"&StationID&"', '"&ProductType&"' , '"&CarPlate&"' , '"&UserID&"' ")

         Rs3.open ("Exec InsertCoupon '"&Barcode&"' , '"&StationID&"' , '"&ProductType&"' , '"&CarPlate&"' , '"&UserID&"' ") ,  conn,3,1
 
         Response.Redirect  "CouponVerification.asp?Message="&iMessage 

         
          Rs3.close
          set Rs3 = nothing

         Else
         
          
          iMessage = "Coupon Existed."          
            
         End if

          Session("SecLevel") = SecLevel

          Session("UserID")   = UserID

          Session("StationID")= StationID

          Response.Redirect  "CouponVerification.asp?Message="&iMessage 
       



%>

