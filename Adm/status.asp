<HTML>
<HEAD>
<TITLE></TITLE>
<META name="description" content="">
<META name="keywords" content="">
<META name="generator" content="CuteHTML">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0000FF" VLINK="#800080">
<%

'========================================================================== 
' 
' VBScript Source File 
' 
' NAME: QueryServer.vbs 
' 
' AUTHOR: Keith Brugnani Webmaster 
' DATE  : 12/30/2013 
' 
' COMMENT: Enumerates all the IIS WWW sites in the given computer's metabase. 
' 
' REVISION: 1.03 
' 
'========================================================================== 
 
'On Error Resume Next 
 
' For seeing how long the script takes to run 
Dim StartTime,EndTime: StartTime = Now  
 
 
WScript.Echo  
Wscript.Echo "StartTime = " & StartTime 
WScript.Echo 
 
Function ShowThisItem(Status) 
   ShowThisItem = False 
   If (Show = ShowAll) Then 
       ShowThisItem = True 
  ElseIf (Show = ShowStarted) And (Status = 2) Then 
       ShowThisItem = True 
  ElseIf (Show = ShowStopped) and (Status = 4) Then 
       ShowThisItem = True 
  ElseIf (Show = ShowPaused) and (Status = 6) Then 
       ShowThisItem = True 
  End If  
 
End Function 
 
Function WebServerState(ServerStatus) 
Select Case ServerStatus 
      Case 1  
         WebServerState = "(starting)" 
      Case 2 
          WebServerState="(started)" 
      Case 3 
          WebServerState="(stopping)" 
      Case 4 
          WebServerState="(stopped)" 
     Case Else 
        WebServerState = "Unknown status : " & ServerStatus 
      End Select 
  End Function 
   
'Uncomment for multiple servers 
'strComputer = Array("MACHINE1", "MACHINE2", "MACHINE3") Add machine name as "MACHINE1", "MACHINE2", "MACHINE3") 
 
'For Localhost or one remote machine 
strComputer = "MACHINE_NAME_HERE" 
 
 
For Each arrComputer In strComputer 'Comment out for a single computer 
 
'Set objWMIService = GetObject("winmgmts:{authenticationLevel=pktPrivacy}\\" & arrComputer & "\root\microsoftiisv2") 
'Set colItems = objWMIService.ExecQuery("Select * from IIsWebServer WHERE Name > 'W3SVC/1'") 'WHERE Name = 'W3SVC/400479726' 
'Uncomment below for debugging 
'WScript.Echo "Server: " & arrComputer & "Server State: " & WebServerState(objItem.ServerState) & " Web Site ID: " & objItem.Name & _ 
' " - " & Err.Number & " - " & Err.Description 
'WScript.Echo "" 
If(Err.Number <> 0) Then 
        ReportError () 
        WScript.Echo "Error trying to open the server: " & arrComputer & vbTab & objItem.Name & " - " & Err.Number & _ 
        Err.Description 
        WScript.Quit (Err.Number) 
    End If 
    WScript.Echo 
    WScript.Echo "*******************************************************************"     
    WScript.Echo "*" 
    WScript.Echo "*" 
    WScript.Echo "*    Web Instance Status On: " & arrComputer 
    WScript.Echo "*" 
    WScript.Echo "*" 
    WScript.Echo "*" 
    WScript.Echo "*******************************************************************" 
Set objParent = GetObject("IIS://" & arrComputer & "/W3SVC") 'Replace arrComputer with strComputer for a single computer 
 
For Each objItem In objParent 
If IsNumeric(objItem.Name) Then 
strOutput = "Site Name: " & objItem.ServerComment & " Site ID: " & objItem.Name 
 
    If (ShowThisItem(objItem.Status) = true) Then 
    Wscript.Echo strOutput & " Server State: " &  WebServerState(objItem.ServerState)  
    End If 
    End If 
    If (Err.Number <> 0) Then 
      ReportError () 
      WScript.Echo "Error trying to query the server: " & objItem.Name & " - " & Err.Number & " - " & Err.Description 
      WScript.Quit (Err.Number) 
    End If 
Next 
Next 'Comment out for a single computer 
 
EndTime = Now 
Wscript.Echo vbCrLf & "EndTime = " & EndTime 
Wscript.Echo "Seconds Elapsed: " & DateDiff("s", StartTime, EndTime) 
Wscript.Echo "Script Complete" 
Wscript.Quit(0) 

%>
</BODY>
</HTML>
