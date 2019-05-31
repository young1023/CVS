<!--#include file="include/CreateJSON.inc" -->
<!--#include file="include/SQLConn.inc" -->
<%

    'initate JSON object
    Dim CouponRespond
    Set CouponRespond = jsObject()
 
     
   ' Retrieve client PC's IP address
   IPAddress = Request.ServerVariables("REMOTE_ADDR")

    ' For testing
   'IPAddress = "127.0.0.1"
   'Response.write IPAddress & "<br>"

      
   ' Retrieve scan date
   ScanDate = Now()
   'Response.write ScanDate & "<br>"

   ' initate record set
   Set Rs0 = Server.CreateObject("Adodb.Recordset")
   Set Rs1 = Server.CreateObject("Adodb.Recordset")
   Set Rs2 = Server.CreateObject("Adodb.Recordset")
   Set Rs3 = Server.CreateObject("Adodb.Recordset")
 
   'Collect Barcode data
   Barcode    = trim(Request("Barcode"))

   'Handle null product type
    ProductType = trim(Request("ProductType"))
    If ProductType = "" Then
       ProductType = "52"
    End If


   'Barcode = "Invalid"
   
   ' ******************************************************************************************************
   '
   ' Check digit 
   '
   ' ******************************************************************************************************

   
    If Barcode = "Invalid" Then

    ' JSON respond when check digit is failed     
     message = "Check digit failed"
     code    = 410
     status  = "Failure"
     
     ' Return JSON
     CouponRespond("program") = "cvs"
     CouponRespond("datetime") = ScanDate
     CouponRespond("Status") =  Status 
     CouponRespond("code") = code
     CouponRespond("message") = message
     CouponRespond.Flush
     Response.end
        
    End IF  

   ' *******************************************************************************************************
   '
   ' Check Coupon Type
   '
   ' *******************************************************************************************************
    Rs0.open ("Exec Check_CouponType '"&Barcode&"', '"&ProductType&"'") ,  conn,3,1

   If Rs0("Tcount") = 0 Then

     ' JSON respond when coupon type is incorrected     
     message = "Coupon Type Incorrect"
     code    = 420
     status  = "Failure"

     
   ' Return JSON
   CouponRespond("program") = "cvs"
   CouponRespond("datetime") = ScanDate
   CouponRespond("Status") =  Status 
   CouponRespond("code") = code
   CouponRespond("message") = message
   CouponRespond.Flush
   Response.end
      
 
    End if

   ' *******************************************************************************************************
   '
   '  Check Barcode Range and Expiry date
   '
   ' *******************************************************************************************************
    Rs1.open ("Exec Checkrange '"&Barcode&"'") ,  conn,3,1

    If Rs1.EoF Then

   ' JSON respond when coupon is not initiated     
     code    = 430
     status  = "Failure"
     message = "Barcode out of range"

     ' Return JSON
    CouponRespond("program") = "cvs"
    CouponRespond("datetime") = ScanDate
    CouponRespond("Status") =  Status 
    CouponRespond("code") = code
    CouponRespond("message") = message
    CouponRespond.Flush
    Response.end
 
   Else


       If Rs1("Expiry_Date") < DateValue(Now) Then

  
         
       ' JSON respond when coupon is expired     
       code    = 440
       status  = "Failure"
       message = "Coupon is expired!"
          
      ' Return JSON
      CouponRespond("program") = "cvs"
      CouponRespond("datetime") = ScanDate
      CouponRespond("Status") =  Status 
      CouponRespond("code") = code
      CouponRespond("message") = message
      CouponRespond.Flush
      Response.end

       End if

   End if

   
     ' Clear Recordset
     Rs0.close
     set Rs0 = nothing
     Rs1.close
     set Rs1 = nothing

   ' *******************************************************************************************************
   '
   '  Check If Barcode existed, else insert coupon
   '
   ' *******************************************************************************************************


    Rs2.open ("Exec CheckCouponExist '"&Barcode&"'") ,  conn,3,1

        ' If there is record
         If Not Rs2.EoF Then

               Present_Date = Replace(Formatdatetime(Rs2("Present_Date")),"/","-")

               StationName  = Rs2("StationName")

            
       ' JSON respond when coupon is existed  
       code    = 450
       status  = "Failure"
       message = "Barcode " & Barcode & " has been scanned on " & Present_Date & " at "& StationName & " station."

       ' Return JSON
       CouponRespond("program") = "cvs"
       CouponRespond("datetime") = ScanDate
       CouponRespond("Status") =  Status 
       CouponRespond("code") = code
       CouponRespond("message") = message
       CouponRespond.Flush
       Response.end

         Else

     Rs3.Open ("Exec InsertCoupon '"&Barcode&"', '"&IPAddress&"','"&ProductType&"','"&ScanDate&"'"), Conn, 3, 1

              ' If record is inserted.
              If Not Rs3.EoF Then 
             
                    ' JSON respond when check digit is successful     
                    code    = 200
                    status  = "success"
                    Message = "Coupon " &  Rs3("Coupon_number") & "was validated successfully on " & scandate

               ' Return JSON
               CouponRespond("program") = "cvs"
               CouponRespond("datetime") = ScanDate
               CouponRespond("Status") =  Status 
               CouponRespond("code") = code
               CouponRespond("message") = message
               CouponRespond.Flush
               Response.end

             Else


                   ' JSON respond when check digit is failed     
                    code    = 460
                    status  = "Failure"
                    Message = "Insert failure."
             
                  ' Return JSON
                  CouponRespond("program") = "cvs"
                  CouponRespond("datetime") = ScanDate
                  CouponRespond("Status") =  Status 
                  CouponRespond("code") = code
                  CouponRespond("message") = message
                  CouponRespond.Flush
                  Response.end

              End if
    
       
      

       
          End if

           ' Clear Recordset
          Rs2.close
          set Rs2 = nothing


    



%>