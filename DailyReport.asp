<!--#include file="include/SQLConn.inc" -->

<%

UserIPAddress = "192.168.1.1"    'Request.ServerVariables("REMOTE_ADDR")

sql = "Select * from station where IPAddress ='" &UserIPAddress& "'"

set rs = conn.execute(sql) 


UserID        = Request.Form("UserID")

If UserID     = "" Then

UserID        = trim(rs("LoginUser"))

End If

'response.write UserID

StationID     = trim(Rs("Station"))

Message       = Request("Message")

Coupon_Number = Request("Coupon_Number")

Car_No        = Request("Car_No")

Face_Value    = Request("Face_Value")

Period        = Request("Period")


' search by date


If Request("SDay") <> "" then

SDay          = Request("SDay")

SMonth        = Request("SMonth") 

SYear         = Request("SYear")


Else

SDay = Day(now())

SMonth = Month(now())

SYear = year(now())

End If

If Request("NDay") <> "" then

NDay          = Request("NDay")

NMonth        = Request("NMonth") 

NYear         = Request("NYear")


Else

NDay = Day(now())

NMonth = Month(now())

NYear = year(now())

End If

Search_Date = SDay & "/" & SMonth & "/" & SYear

'Search_Date =  CDate(Search_Date)

Search_NDate = NDay & "/" & NMonth & "/" & NYear

'response.write UserID

%>

<HTML>
<HEAD>
<title>
    --  禮券驗證系統 -- 
</title>


<meta http-equiv="Content-Language" content="zh-hk">
<meta http-equiv="Content-Type" content="text/html; charset=big5" /> 
<link rel="stylesheet" type="text/css" media="all" href="include/publish.css" />
<link rel="shortcut icon" href="images/favicon.ico" />
<!--#include file="include/JavaScript.js" -->
</HEAD>
<body class="homepage" onload="document.fm1.Barcode.focus();">



<!--#include file="include/header.inc" -->



<form name="fm1" method="post" action="">

<div align="center">

<h2 class="Title">禮券驗證系統</h2>

<span class="noprint">
<div align="center"><a href="CouponVerification.asp">驗證</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="logout.asp">登出</a></div>
</span>


 <table width="98%" border="1" cellpadding="4" cellspacing="1" class="Report">
                      
                        


         
			  <tr>
							
                                <td height="28"> 
                                  <%

		pageid=trim(request.form("pageid"))

		if pageid="" then
		  pageid=1
		end if

        Barcode = replace(trim(request.form("Barcode")),"%","％")

        Barcode = replace(Barcode,"'","''")

       
        Face_Value = Left(Barcode, 3)

        Coupon_Type = Mid(Barcode, 5,2)

        Coupon_batch = Mid(Barcode,7,3)

        Coupon_Number = Mid(Barcode, 10,6)

       
        

   ' Start the Queries    
   ' *****************
      
   fsql = "SELECT m.Product_Type as ProductType, * FROM MasterCoupon m INNER JOIN CouponRequest c "

   fsql = fsql & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

   fsql = fsql & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

   fsql = fsql & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID


  ' Search Coupon Number
  ' ********************
   if Barcode <> "" then

       If len(Barcode) > 14 then

       fsql = fsql & " and Coupon_Number = '"& Coupon_Number &"'"

       fsql = fsql & " and Face_Value = '"& Face_Value &"'"

       fsql = fsql & " and Coupon_Type = '"& Coupon_Type &"'"

       fsql = fsql & " and Coupon_Batch = '"& Coupon_Batch &"'"

       Else

       fsql = fsql & " and Coupon_Number LIKE '%"&Barcode&"%' " 

       End If

   end if


 
  ' Search by Date
  ' **************


     fsql = fsql & " and  Present_Date >=   DATEADD(hh, -2, Convert(datetime, '" & Search_Date &"', 105)) " 

  
     fsql = fsql & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 


   ' By UserID

    if UserID <> "All"  then


      fsql = fsql & " and Period = '"& UserID &"' " 

    end if
   
 
       fsql = fsql & " order by present_date"

        'response.write fsql
        'response.end

        set frs=createobject("adodb.recordset")
		frs.cursortype=1
		frs.locktype=1
        frs.open fsql,conn

%>

<span class="noprint">

<%

        if frs.RecordCount=0 then
           response.write "<font color=red>沒有紀錄</font>"
           'response.end
        else


          findrecord=frs.recordcount
          response.write "共有 <font color=red>"&findrecord&"</font> 條紀錄 ;"

  
         frs.PageSize = 200

         call countpage(frs.PageCount,pageid)

         end if

%>


	     &nbsp;&nbsp;<input type="text" name="Barcode" size="15" maxlength=15 value="">
		 &nbsp;&nbsp;<input type="button" value="禮券編號" onClick="findenum();">

</span>	   


                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 


   <table border="1" align=center cellpadding="4" cellspacing="1" class="DailyReport" width="100%" height="100%">

<tr bgcolor="#DFDFDF">
          <td colspan="8" align="right">油站</td>
          <td align="center"><% = StationID %></td>
      </tr>
    <tr> 
   

			<td colspan="3"> 日期:
			 
			<select name="SDay" class="common" class="reportdate">

        <% 
              j = 1

            for i = 1 to 31 

         %>

	<option value="<%=j%>" <% if trim(SDay) = trim(j) Then response.write "selected" %>><% = j %></option>
		
     <%
     
             j = j + 1

             Next 

         %>
		
			
			</select>


			<select name="SMonth" class="common"> 
 
        <% 
              k = 1

            for l = 1 to 12 

         %>
          	
    <option value="<%=k%>" <% if trim(SMonth) = trim(k) then response.write "selected"%>><% = Left(MonthName(k),3) %></option>
			 <%
     
             k = k + 1

             Next 

         %>

</select>


			<select name="SYear" class="common">   
<% 
Dim Year_starting
Dim Year_ending

Year_starting = Year(DateAdd("yyyy", -1, Now()))
year_ending = Year(Now())

for m = Year_starting to Year_ending
%>			         
			<option value="<%=m%>" <% if clng(m)=clng(SYear) then response.write "selected"%>><%=m%></option>

<% next %>

			</select> 
<span class="reportdate">

<% = SDay & " " & Left(MonthName(SMonth),3) & " " & SYear %>

</div>
			</td>


	<td colspan="3">
			 
			<select name="NDay" class="common">

        <% 
              j = 1

            for i = 1 to 31 

         %>

	<option value="<%=j%>" <% if trim(NDay) = trim(j) Then response.write "selected" %>><% = j %></option>
		
     <%
     
             j = j + 1

             Next 

         %>
		
			
			</select>


			<select name="NMonth" class="common"> 
 
        <% 
              k = 1

            for l = 1 to 12 

         %>
          	
    <option value="<%=k%>" <% if trim(NMonth) = trim(k) then response.write "selected"%>><% = Left(MonthName(k),3) %></option>
			 <%
     
             k = k + 1

             Next 

         %>

</select>


			<select name="NYear" class="common">   
<% 


Year_starting = Year(DateAdd("yyyy", -1, Now()))

for i = Year_starting to Year(Now())
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(NYear) then response.write "selected"%>><%=i%></option>

<% next %>

			</select> 

	<select name="NHour" class="common">   
<% 


Year_starting = Year(DateAdd("yyyy", -1, Now()))

for i = 1 to 24
%>			         
			<option value="<%=i%>" <% if clng(i)=clng(NHour) then response.write "selected"%>><%=i%></option>

<% next %>

			</select> 

<span class="reportdate">

<% = NDay & " " & Left(MonthName(NMonth),3) & " " & NYear %>

</div>
			</td>

<td colspan="2">更期:

            <select size="1" name="UserID" class="common">
            <option value="All" <% If UserID = "All" then%>Selected<%End If%>>全日</option>
			<option value="11" <% If UserID = "11" then%>Selected<%End If%>>早更</option>
			<option value="12" <% If UserID = "12" then%>Selected<%End If%>>中更</option>
            <option value="13" <% If UserID = "13" then%>Selected<%End If%>>晚更</option>
	</select>

</td>

	<td ><input type="button" value="更結報告" onClick="Report();" class="noborder">
	</td>
    
</tr>


<% If i = 25 then %>
<tr bgcolor="#DFDFDF" class="page-break">
<% Else%>
<tr bgcolor="#DFDFDF">
<%End if%>
   

<td width="15%">日期 / 時間</td>
<td width="10%">更期</td>
<td width="10%">銀碼</td>
<td width="10%">類型</td>
<td width="10%">批次</td>
<td width="15%">禮券編號</td>
<td width="10%">產品</td>
<td width="10%">機號</td>
<td width="10%">電子禮券</td>
</tr>

<%

Total = 0
Total2 = 0
Total3 = 0

i=0

 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
   
  
%>


   <tr>

<td width="30%">
<% = i & ". " & frs("Present_Date") %>
</td>

<td>
<%

Period = frs("Period")

If Period = 11 Then

   Response.write "早"

Elseif Period = 12 Then

   Response.write "中"

Elseif Period = 13 Then

   Response.write "晚"

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
<% = frs("ProductType") %>
</td>

<td>
<% = frs("MachineNo") %>
</td>

<td>
<% = frs("Digital") %>
</td>

</tr>



<%

   
   frs.movenext
  loop
 end if
  %>
<tr>
<td colspan = "2" align="right">
Total:
</td>
<td><% = Total %></td>
<td colspan= "7" ></td>
</tr>

                                  </table>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 
<script language=JavaScript>
<!--

function gtpage(what)
{
document.fm1.pageid.value=what;
document.fm1.action="DailyReport.asp"
document.fm1.submit();
}

function findenum()
{
document.fm1.action="DailyReport.asp"
document.fm1.submit();
}

function Report()
{
document.fm1.action="DailyReport.asp"
document.fm1.submit();
}

//-->
</script>

<span class="noprint">
<%
	 if frs.recordcount>0 then
             call countpage(frs.PageCount,pageid)
			 'response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid)<>1 then
                 'response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 'response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 else
                 'response.write " First "
                 'response.write " Previous "
			 end if
	         if Clng(pageid)<>Clng(frs.PageCount) then
                 'response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 'response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "
             else
                 'response.write " Next "
                 'response.write " Last "
			 end if
	         'response.write "&nbsp;&nbsp;"
	 end if
%>
</span>
                                </td>
                              </tr>


<tr>

    <td>


<%

   ' Start the Queries for eCoupon  
   ' *****************************
      
     sql = "SELECT MachineNo, Digital,  Face_Value,  Count(Face_Value) as eCount, Sum(Cast(Face_Value as float)) as eAmount FROM MasterCoupon m INNER JOIN CouponRequest c "

     sql = sql & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

     sql = sql & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

     sql = sql & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID

     sql = sql & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

     sql = sql & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 

   ' By UserID

    if UserID <> "All"  then

      sql = sql & " and Period = '"& UserID &"' " 

    end if 
 
      sql = sql & " Group by MachineNo, Digital, Face_Value "

      'response.write sql


     Set rs = Conn.Execute(sql)


  %>

         <table width="60%" border="1" cellpadding="4" cellspacing="0" class="DailyReport">
 
<tr bgcolor="#DFDFDF">
  
          <td colspan="5">(總結 (以銀碼分類))</td>

     </tr>

             <tr>

                    <td width="20%">電腦編號</td>

                    <td>電子禮券</td>

                    <td>銀碼</td>

                    <td>數量</td>

                    <td>總值</td>

             </tr>

<%

    If Not rs.EoF Then

       rs.MoveFirst

     Do While Not rs.Eof
 
  %>           
             <tr>

                    <td><% = rs("MachineNo") %></td>

                    <td><% = rs("Digital") %></td>

                    <td><% = rs("Face_Value") %></td>

                    <td><% = rs("eCount") %></td>

                    <td><% = rs("eAmount") %>

<%  

Total2 = Total2 + rs("eAmount") 

%>

                     </td>


             </tr>
<%
            rs.MoveNext

            Loop

   End If

%>

     
      <tr>
                    <td colspan="3"></td>
                    <td >總數</td>
                    <td><% = Total2 %></td>

  </tr>   

         </table>

    </td>

</tr>

<tr>

    <td>


<%

   ' Start the Queries for eCoupon for Batch
   ' ****************************************
      
     sql2 = "SELECT MachineNo, Digital, Coupon_Type, Batch, Count(Batch) as eCount2, Sum(Cast(Face_Value as float)) as eAmount2 FROM MasterCoupon m INNER JOIN CouponRequest c "

     sql2 = sql2 & "ON m.Coupon_Type = c.Product_Type AND m.Coupon_Batch = c.Batch  "

     sql2 = sql2 & "AND cast(m.Face_Value as decimal(9,0))   = c.FaceValue AND m.Coupon_Number <="

     sql2 = sql2 & "c.End_Range AND m.Coupon_Number >= c.Start_Range where m.RequestedID = "&StationID

     sql2 = sql2 & " and  Present_Date >=   Convert(datetime, '" & Search_Date &"', 105) " 

     sql2 = sql2 & " and  Present_Date < DATEADD(dd,DATEDIFF(dd,0, Convert(datetime, '" & Search_NDate &"', 105)),0) + 1 " 

   ' By UserID

    if UserID <> "All"  then

      sql2 = sql2 & " and Period = '"& UserID &"' " 

    end if 
 
      sql2 = sql2 & " Group by MachineNo, Coupon_Type, Digital, Batch, Face_Value"

      'response.write sql2


     Set rs2 = Conn.Execute(sql2)


  %>

         <table width="60%" border="1" cellpadding="4" cellspacing="0" class="DailyReport">
 
<tr bgcolor="#DFDFDF">
  
          <td colspan="6">總結 (以批次分類)</td>

     </tr>


             <tr>

                    <td width="20%">電腦編號</td>

                    <td>電子禮券</td>

                    <td>類型</td>

                    <td>批次</td>

                    <td>數量</td>

                    <td>總值</td>



             </tr>

<%

    If Not rs2.EoF Then

       rs2.MoveFirst

     Do While Not rs2.Eof
 
  %>           
             <tr>

                    <td><% = rs2("MachineNo") %></td>

                    <td><% = rs2("Digital") %></td>

                    <td><% = rs2("Coupon_Type") %></td>

                    <td><% = rs2("Batch") %></td>

                    <td><% = rs2("eCount2") %></td>

                    <td><% = rs2("eAmount2") %>
<%  

Total3 = Total3 + rs2("eAmount2") 

%>


</td>


 
             </tr>
<%
            rs2.MoveNext

            Loop

   End If

%>

      <tr>
                    <td colspan="5"></td>
                    <td><% = Total3 %></td>

  </tr>   
        

         </table>

    </td>

</tr>

                              <tr> 
                                <td height="28" align="center">
<span class="noprint"> 
<input type="button" value="   打印   "  onClick="window.print();" class='common'>

<input type="button" value="   Excel   "  onClick="window.doConvert();" class='common'>

</span>

<%

              response.write "<input type=hidden value="&pageid&" name=pageid>"
			  frs.close
			  set frs=nothing
			  conn.close
			  set conn=nothing
%>
                                </td>
                              </tr>
                            </table>
                          </form>

                          


<%

  ' function
  Sub countpage(PageCount,pageid)
  response.write "第 "&pagecount&"</font> 頁 "
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub


%>
<SCRIPT language=JavaScript>
<!--
function doConvert(){
window.open("ConvertReport1.asp?Station=<%=StationID%>&Sday=<%=SDay%>&SMonth=<%=SMonth%>&SYear=<%=SYear%>&NDay=<%=NDay%>&NMonth=<%=NMonth%>&NYear=<%=NYear%>&UserID=<%=UserID%>"); 

}

//-->
</SCRIPT>
</body>
</HTML>
