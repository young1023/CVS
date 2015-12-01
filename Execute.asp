 <!--#include file="include/SQLConn.inc" -->
<%


   Set Rs1 = Server.CreateObject("Adodb.Recordset")
   Set Rs2 = Server.CreateObject("Adodb.Recordset")
   Set Rs3 = Server.CreateObject("Adodb.Recordset")
 
   StationID = Request("StationID")

   Barcode    = Request("Barcode")

   ProductType = Request("ProductType")

   CarPlate   = Request("CarPlate")

  

   response.write "Station ID:  " & Stationid & "<br>"

   response.write "Barcode: " & Barcode & "<br>"

   'response.end

   Set Message = Nothing

    Rs1.open ("Exec Checkrange '"&Barcode&"'") ,  conn,3,1

   Response.write ("Exec Checkrange '"&Barcode&"'")

   'Response.end

   If Rs1.EoF Then

           Message = "驗證失敗 - 禮券無效 "


    Response.Redirect  "CouponVerification2.asp?&ProductType="&ProductType&"&Color=2&Message="&Message
      
 
   Else


       If Rs1("Expiry_Date") < DateValue(Now) Then

  
          Message = "驗證失敗 - 禮券過期!"
        
          

    Response.Redirect  "CouponVerification2.asp?&ProductType="&ProductType&"&Color=2&Message="&Message
      

       End if

   End if


          Rs1.close
          set Rs1 = nothing

          Rs2.open ("Exec CheckCouponExist '"&Barcode&"'") ,  conn,3,1

          Response.write ("Exec CheckCouponExist '"&Barcode&"'")

     
         If Not Rs2.EoF Then

         Present_Date = Rs2("Present_Date")

         RequestedID  = Rs2("RequestedID")

         StationName  = Rs2("StationName")

         'response.write Present_Date
         'response.end
         Color   = "2"

         Message = "驗證失敗 - 重複使用! 禮券 " & Barcode & " 曾於 " & Present_Date & " 在"& StationName & "站驗證!"
         
       
         Else

        
          Message = "驗證成功!"

          Color   = "1"

     

          Rs3.Open ("Exec InsertCoupon '"&Barcode&"', '"&StationID&"','"&ProductType&"','"&CarPlate&"'"), Conn, 3, 1


          End if


         Response.Redirect  "CouponVerification2.asp?&ProductType="&ProductType&"&Color="&Color&"&Message="&Message
      

          Rs2.close
          set Rs2 = nothing

 
%>
     