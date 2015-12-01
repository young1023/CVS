<!--#include file="include/SQLConn.inc" -->
<%

StationID     = Request("Station"))

' Tells the browser to open excel
Response.ContentType = "application/vnd.ms-excel" 
Response.addHeader "content-disposition","attachment;filename=DailyReport_Station"&StationID&"_"&Month(Now())&Year(now())&".xls"

' Received parameters 

SDay          = Request("SDay")

SMonth        = Request("SMonth") 

SYear         = Request("SYear")

NDay          = Request("NDay")

NMonth        = Request("NMonth") 

NYear         = Request("NYear")

UserID        = Request("UserID")

Search_Date = SDay & "/" & SMonth & "/" & SYear

Search_NDate = NDay & "/" & NMonth & "/" & NYear


%>

<HTML>
<HEAD>
<title>
    --  §�����Ҩt�� -- 
</title>


<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" /> 
</HEAD>
<body>
<%


' Start the Queries
' *****************
      
   fsql = "SELECT m.Product_Type as ProductType, * FROM MasterCoupon m INNER JOIN CouponRequest c "

   fsql = fsql & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

   fsql = fsql & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

   fsql = fsql & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID

 
  ' Search by Date
  ' **************


    fsql = fsql & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

  
    fsql = fsql & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 

    
    if UserID <> "" then

    fsql = fsql & " and Period = '"& UserID &"' " 

    end if

    fsql = fsql & " order by id desc"

    set frs= conn.execute(fsql)
	
%>


<table>
<tr bgcolor="#DFDFDF">
<td width="15%">��� / �ɶ�</td>
<td width="10%">���</td>
<td width="10%">�ȽX</td>
<td width="10%">����</td>
<td width="10%">�妸</td>
<td width="15%">§��s��</td>
<td width="10%">���~</td>
<td width="10%">����</td>
<td width="10%">�q�l§��</td>
</tr>

<%

  do while not frs.eof

   if frs.eof then exit do
   
  
%>

<tr>
<td width="30%">
<% = frs("Present_Date") %>
</td>

<td>
<%
If frs("Period") = 11 Then

   Response.write "��"

Elseif frs("Period") = 12 Then

   Response.write "��"

Elseif frs("Period") = 13 Then

   Response.write "��"

End if
%>
</td>

<td>
<%= frs("Face_Value") %>
<%  

Total = Total + frs("Face_Value") 

%>
</td>

<td>
 <% = frs("Coupon_type") %>
</td>

<td>
 <% = frs("Coupon_Batch") %>
</td>

<td>
<% = frs("Coupon_Number") %>
</td>

<td>
<% = frs("Product_Type") %>
</td>

<td>

</td>

<td>
<% = frs("Digital") %>
</td>

</tr>



<%

   
   frs.movenext

  loop
 

  %>

</table>
                                
</body>
</HTML>
