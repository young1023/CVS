<% 
'on error resume next
'*********************************************************************************
'NAME       : Report.asp          
'DESCRIPTION: Daily Report
'INPUT      : 
'OUTPUT     : 
'RETURNS    :                     
'CALLS      :                     
'CREATED    : 131017 Gary Yeung
'MODIFIED   : 
'********************************************************************************
pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
%>


<HTML>
<HEAD>
	<link rel="stylesheet" type="text/css" href="include/uob.css" />

<TITLE>更結報告</TITLE>

<!-- Load the javascript code -->

<SCRIPT language=JavaScript>
<!--

function AssignValue(){
	
  	myString = this.myform.myselect.value;
 	  self.opener.document.fm1.ClientTo.value=myString;
 	  self.opener.document.fm1.ClientFrom.value=myString;
 		
 		self.close();
 	
}


function gtpage(what)
{
document.myform.pageid.value=what;
document.myform.action="SearchClientNumber.asp?sid=<%=SessionID%>"
document.myform.submit();
}


//-->
</script>


</HEAD>

<BODY OnLoad="document.myform.keyword.focus();document.myform.keyword.select();">

								<FORM NAME="myform"  method="post"  action="<%= strURL %>?sid=<%=SessionID%>">
								  <table width="99%" border="0" class="normal">
			
							
										<tr><td class="common"> 
											
											Enter client No., English Name or Chinese Name
										</td></tr>
										<tr><td class="common"> 
										<INPUT name="keyword" value="<%= Search_keyword %>">
											<INPUT TYPE=submit value="Filter">
											<INPUT TYPE=Reset value="Clear">

										</td></tr>

			

								
<%							
		
                Set Rs = server.createobject("adodb.recordset")  

       
              
                sql = "Select * from Client c Left Join Member m on c.AECode = m.LoginName Where 1=1 "

                
                If Search_AECode <> "" Then

                sql = sql & "and c.AECode in ('"&Search_AECode&"') "

                End If


                If Search_AEGroup <> "" Then

                sql = sql & "and m.GroupID = '"&Search_AEGroup&"' "

                End If

       
                If Search_keyword <> "" Then
                
                sql = sql & "and (ClntCode like '%"&Search_keyword&"%' or c.name like  '%"&Search_keyword&"%' or cname like '%"&Search_keyword&"%' or Ename like '%"&Search_keyword&"%' )"

                End If

                sql = sql & "Order by ClntCode"
                
                'Response.write sql
                
            
       		Set Rs1 = server.createobject("adodb.recordset")
			    'Rs1.open ("exec Retrieve_ClientNumber '"&Search_AECode&"', '"&Search_AEGroup&"', N'"&Search_keyword&"', '"&Search_SharedGroup&"' ") ,  StrCnn,3,1
			'response.write ("exec Retrieve_ClientNumber '"&Search_AECode&"', '"&Search_AEGroup&"', N'"&Search_keyword&"', '"&Search_SharedGroup&"'  ")
			
			
			  
		         Rs1.cursortype=3
		         Rs1.locktype=1
                 Rs1.open sql,conn

			if Rs1.RecordCount=0 then

				
					'no record found
					response.write ("<tr><td>No record found</td></tr>")
				
				else

			  	'Set rs1 = rs1.NextRecordset() 
	
         Rs1.PageSize = 10
         
         
	 
	
	%>		 
								<tr><td class="common">
								Select client (Total <%=Rs1.RecordCount%> clients), 
 <%
	 if Rs1.recordcount>0 then
             
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid) <>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 end if
             call countpage(Rs1.PageCount,pageid)
	         if Clng(pageid)<>Clng(Rs1.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&Rs1.PageCount&"') style='cursor:hand' >Last</a> "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if
%>
								</td></tr>
							
								
								<tr><td class="common">
								
								<SELECT NAME="myselect" SIZE=10 class="common">
									<% 'do while (  Not rs1.EOF) %>
									
									<%
									
		

									
 i=0
 if Rs1.recordcount>0 then
  Rs1.AbsolutePage = pageid
  do while (Rs1.PageSize-i)
   if Rs1.eof then exit do
   i=i+1

  
%>

												<Option value="<%=rs1("clntcode")%>"> <% response.write Rs1("clntcode") + " : " +  rs1("ename") + " " +  rs1("cname")%> </OPTION>
			
															
									<%
										rs1.movenext 
										loop 
									%>
								</SELECT>
								
								
								
								<SCRIPT TYPE="text/javascript">
								<!--
								myClientNumber = ""
								var myfilter = new filterlist(document.myform.myselect);
								//-->
								</SCRIPT>
							</td></tr>
							
							<tr><td class="common">
							<INPUT TYPE=button onClick="AssignValue();" value="Select Client">
							</td></tr>
			<%
			
			     response.write "<input type=hidden value="&pageid&" name=pageid>"

					'record found end if
					end if
						
			'argument end if						
			end if 
			%>
			
			
								
								
								</table>
								</FORM>
							

					

			<% 

		end if
			%>
	<%
	'*******
	' END  
	'*******
	%>
	

</BODY>
</HTML>