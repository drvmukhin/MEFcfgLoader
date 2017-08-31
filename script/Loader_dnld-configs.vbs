#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'		JUNIPER MEF CONFIG DOWNLOAD SCRIPT
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 75
Dim nResult
Dim strLine
Dim nOverwrite
Dim strMonthMaxFileName, strFileString, strSkip, strFileButton, strFileInventory, strFileSession
Dim strDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke
Dim strDeviceID, strAccountID
Dim nDebug, nInfo
Dim nIndex, nInd, nCount
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile
Dim vSession(30), vSettings
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand
Dim Platform
Dim objTab_L, objTab_R
Dim vWaitForCommit, vModels, vWaitForFtp
vWaitForftp = Array("No route to host","Connected to","Connection refused")
vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
vModels = Array("acx5096","acx5048","acx1100","acx1000","acx2100","acx2200","mx80","mx104","mx240","mx480","mx960")
Dim strFileSettings,strFileParam
Dim vDelim, vParamNames
    Const UNKNOWN = "Unknown"
	Const SECURECRT_FOLDER = "SecureCRT Folder"
	Const SECURECRT_SESSION_FOLDER = "SecureCRT Session Folder"
	Const FTP_SERVER_FOLDER = "FTP Server Folder"
    Const WORK_FOLDER = "Work Folder"
    Const CONFIGS_FOLDER = "Configuration Files Folder"
    Const CONFIGS_PARAM  = "MEF Service Parameters"
    Const CONFIGS_GLOBAL  = "CONFIGS_GLOBAL"
    Const CONFIGS_RE0  = "CONFIGS_RE0"
    Const CONFIGS_RE1  = "CONFIGS_RE1"
    Const Node_Left_IP  = "Left Node IP"
    Const Node_Right_IP  = "Right Node IP"
	Const HIDE_CRT = "Hide Terminal Session"
    Const FTP_IP  = "FTP IP"
    Const FTP_User  = "FTP User"
    Const FTP_Password  = "FTP Password"
	Const UNI_A = "UNI-A"
	Const UNI_B = "UNI-B"
	Const UNI_C = "UNI-C"
	Const UNI_D = "UNI-D"
	Const UNI_CC = "UNI-CC"
	Const UNI_DD = "UNI-DD"
	Const NNI = "NNI"
	Const PLATFORM_NAME = "Platform Name"
	Const PLATFORM_INDEX = "Node Name Prefix"
	Const Template = "XLS TEMPLATE"
	Const Orig_Folder = "Original TCG Templates"
	Const Dest_Folder = "Exported TCG Templates"
	Const WorkBookPrefix = "WorkBookPrefix"
	Const SECURECRT_L_SESSION = "Left Node Session"
	Const SECURECRT_R_SESSION = "Right Node Session"


ReDim vSettings(30)
ReDim vParamNames(UBound(vSettings))
    vParamNames(0) = Node_Left_IP
    vParamNames(1) = Node_Right_IP
    vParamNames(2) = FTP_IP
    vParamNames(3) = FTP_User
    vParamNames(4) = FTP_Password
	vParamNames(5) = SECURECRT_FOLDER
    vParamNames(6) = WORK_FOLDER
    vParamNames(7) = CONFIGS_FOLDER
    vParamNames(8) = Orig_Folder
	vParamNames(9) = Dest_Folder
	vParamNames(10) = SECURECRT_L_SESSION
	vParamNames(11) = SECURECRT_R_SESSION
    vParamNames(12) = CONFIGS_PARAM
	vParamNames(13) = PLATFORM_NAME
    vParamNames(14) = PLATFORM_INDEX
    vParamNames(15) = SECURECRT_SESSION_FOLDER
    vParamNames(16) = FTP_SERVER_FOLDER
    vParamNames(17) = HIDE_CRT
    vParamNames(18) = UNKNOWN
    vParamNames(19) = UNKNOWN
	vParamNames(20) = UNKNOWN
    vParamNames(21) = UNKNOWN
    vParamNames(22) = UNKNOWN
    vParamNames(23) = UNKNOWN
	vParamNames(24) = UNKNOWN
	vParamNames(25) = UNKNOWN
	vParamNames(26) = UNKNOWN
    vParamNames(27) = CONFIGS_GLOBAL
	vParamNames(28) = CONFIGS_RE0
    vParamNames(29) = CONFIGS_RE1

	vDelim = Array("=",",",":")	
nDebug = 0
nInfo = 1
Platform = "acx"
strVersion = "None"
strFileSettings = "settings.dat"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = CreateObject("WScript.Shell")
Sub Main()
'------------------------------------------------------------------
'	CHECK NUMBER OF ARGUMENTS AND EXIT IF LESS THEN 3
'------------------------------------------------------------------
	If crt.Arguments.Count < 5 Then
			MsgBox "ERROR: Wrong number of arguments" & chr(13) &_
			"ARG1: Service Type" & chr(13) &_
			"ARG2: Service Flavour" & chr(13) &_
			"ARG3: Test Number#" & chr(13) &_
			"ARG4: Full Path to settings.dat file" & chr(13) &_
			"ARG5: Full Path for Loader Work Folder"
		crt.quit
		Exit Sub
	End If
	strService = crt.Arguments(0)
	strFlavor = crt.Arguments(1)
	strTask = crt.Arguments(2)
	strFileSettings = crt.Arguments(3)
	strDirectoryWork = crt.Arguments(4)		
'----------------------------------------------------------------
'	Open log File
'----------------------------------------------------------------
			n = 5
			i = 0
			nRetries = 5
				Do While i < nRetries
					On Error Resume Next
					Err.Clear
						Set objDebug = objFSO.OpenTextFile(strDirectoryWork & "\Log\" & "debug-terminal.log",ForWriting,True)
						Select Case Err.Number
							Case 0
								Exit Do
							Case 70
								i =  i + 1
								n = 3
								crt.sleep 100 * n
							Case Else 
								Exit Do		
						End Select
				Loop
				On Error goto 0
'--------------------------------------------------------------------
'   LOOKING FOR EXISTED MONITOR SESSION (tail.exe)
'--------------------------------------------------------------------
strLaunch = strDirectoryWork & "\bin\tail.exe -f " & strDirectoryWork & "\log\debug-terminal.log"
If Not GetAppPID(strPID, strParentPID, "tail.exe") Then 
    objEnvar.run (strLaunch)
Else
    Call FocusToParentWindow(strPID)
End If
Call TrDebug_No_Date ("GetMyPID: PID = " & strPID & " ParentPID = " & strParentPID,"",objDebug, MAX_LEN, 1, nDebug)								
	'Set original values for the settings values
	Call SetOriginalSettings(vSettings)		
    'Load Settings
    Select Case LoadSettings(strFileSettings, strFileParam)
	    Case -1
			MsgBox "Cant' find Settings file: " & chr(13) & strFileParam
			Exit Sub
		Case -2 
			MsgBox "Cant' find Settings file: " & chr(13) & strFileSettings
			Exit Sub
	End Select
	'-------------------------------------------------------------------------------------------
	'  	LOAD INITIAL CONFIGURATION FROM SETTINGS FILE
	'-------------------------------------------------------------------------------------------
	strDirectoryVandyke = GetValue(SECURECRT_FOLDER)
	strDirectoryConfig = GetValue(7)
	strCfgGlobal = GetValue(27)
	strCfgRE0 =    GetValue(28)
	strCfgRE1 =    GetValue(29)
	DUT_Platform = GetValue(13)
	Platform =     GetValue(14)
	strLeft_ip =   GetValue(0)
	strRight_ip =  GetValue(1)
	strFTP_ip =    GetValue(2)
	strFTP_name =  GetValue(3)
	strFTP_pass =  GetValue(4)
	vWaitForftp(1) = "Connected to " & strFTP_ip
	'--------------------------------------------------------------------------------
	'          GET NAME OF THE TELNET SESSIONS
	'--------------------------------------------------------------------------------
	strLine = GetValue(10)
	strFolder =   Trim(Split(strLine, ",")(0)) & "/"
	strSessionL = Trim(Split(strLine, ",")(1))
	strHostL =    Trim(Split(strLine, ",")(2)) 
	strLogin =    Trim(Split(strLine, ",")(3))

	strLine = GetValue(11)
	strFolder =   Trim(Split(strLine, ",")(0)) & "/"
	strSessionR = Trim(Split(strLine, ",")(1))
	strHostR =    Trim(Split(strLine, ",")(2)) 
	strLogin =    Trim(Split(strLine, ",")(3))
	'-------------------------------------------------------------------------------------------
	'  	LOAD TESTBED TOPOLOGY
	'-------------------------------------------------------------------------------------------
	If Not objFSO.FileExists(strFileParam) Then 
		MsgBox "MEF Parameters File: " & chr(13) & strFileParam  & " not found!"
		Exit Sub
	End If
'------------------------------------------------------------------
'	INITIAL CONFIGURATION FILES
'------------------------------------------------------------------
	strGlobalFileL = strCfgGlobal & "-" & Platform & "-l.conf"
	strGlobalFileR = strCfgGlobal & "-" & Platform & "-r.conf"
	strRe0FileL = strCfgRE0 & "-" & Platform & "-l.conf"
	strRe0FileR = strCfgRE0 & "-" & Platform & "-r.conf"
'------------------------------------------------------------------
'	LOG MAIN VARIABLES
'------------------------------------------------------------------
	Call TrDebug_No_Date ("TelnetScript: strSessionL, strHostL: " & strFolder & strSessionL & ", " & strHostL,"", objDebug, MAX_LEN, 1, nDebug)						
	Call TrDebug_No_Date ("TelnetScript: " & strGlobalFileL & ", " & strGlobalFileR,"", objDebug, MAX_LEN, 1, nDebug)						
	Call TrDebug_No_Date ("TelnetScript: " & strRe0FileL & ", " & strRe0FileR, "", objDebug, MAX_LEN, 1, nDebug)						
'------------------------------------------------------------------
'	START MAIN PROGRAM
'------------------------------------------------------------------
	strDirectoryConfig = strDirectoryConfig & "\" & strService
	strConfigFileL = strFlavor & "-" & strTask & "-" & Platform & "-l.conf"
	strConfigFileR = strFlavor & "-" & strTask & "-" & Platform & "-r.conf"
	'--------------------------------------------------------------------------------
    '  Start SSH session to Left Node
    '--------------------------------------------------------------------------------
	Call TrDebug_No_Date ("START DOWNLOADING CONFIGURATION FILES ","", objDebug, MAX_LEN, 3, nInfo)						
	On Error Resume Next
	Err.Clear
	Set objTab_L = crt.Session.ConnectInTab("/S " & strFolder & strSessionL)
	If Err.Number <> 0 Then 
		Call  TrDebug_No_Date ("ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, nInfo)
		Exit Sub
	End If
	On Error Goto 0
	objTab_L.Caption = strHostL
	objTab_L.Screen.Synchronous = True
	objTab_L.Screen.Send chr(13)
	nResult = objTab_L.Screen.WaitForString ("@" & strHostL & ">",2)
    If nResult = 0  Then
		If IsObject(objDebug) Then Call  TrDebug_No_Date (strHostL & " ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, nInfo) End If
		objTab_L.Session.Disconnect
    	Exit Sub
	End If
'---------------------------------------------
'   SAVE main config L
'---------------------------------------------
	objTab_L.Screen.Send "edit" & chr(13)
	objTab_L.Screen.WaitForString "@" & strHostL & "#"
	objTab_L.Screen.Send "save " & strConfigFileL & chr(13)
	objTab_L.Screen.WaitForString "@" & strHostL & "#"	
	objTab_L.Screen.Send "exit" & chr(13)
	objTab_L.Screen.WaitForString "@" & strHostL & ">"	
	Call TrDebug_No_Date ("L NODE: SAVING CONFIGURATION LOCALLY: " & strConfigFileL ,"OK", objDebug, MAX_LEN, 1, nInfo)						
'---------------------------------------------
'   FTP CONNECTING
'---------------------------------------------	
	objTab_L.Screen.Send "ftp "  & strFTP_ip & chr(13)
	nResult = objTab_L.Screen.WaitForStrings (vWaitForFtp, 5)
     Select Case nResult
        Case 0
			Call  TrDebug_No_Date ("HOST " & strHostL, ": FTP CONNECTION TIME OUT", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("CONFIGURATION UPLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)					
'           objTab_L.Session.Disconnect			
'			crt.quit
			Exit Sub
        Case 1 
			Call  TrDebug_No_Date ("HOST " & strHostL & " HAS NO ROUTE TO FTP SERVER", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("1. Check FTP Server IP-address under MEF Loader Settings ", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("2. Make sure the Left node has a route to FTP server IP address ","", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("CONFIGURATION UPLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)		
	        objTab_L.Session.Disconnect			
    		crt.quit
    		Exit Sub
        Case 2 
			Call  TrDebug_No_Date ( "CONNECTING TO FTP", "OK", objDebug, MAX_LEN, 1, nInfo)   
		Case 3
			Call  TrDebug_No_Date ("FTP CONNECTION REFUSED", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("1. Make sure that FTP Server is running ", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("2. Make sure connection is not blocked by firewall","", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("CONFIGURATION UPLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)		
	        objTab_L.Session.Disconnect			
    		crt.quit
    		Exit Sub	
     End Select		
	objTab_L.Screen.WaitForString "Name (" & strFTP_ip & ":" & strLogin & "):"
	objTab_L.Screen.Send strFTP_name & chr(13)
	objTab_L.Screen.WaitForString "Password:"
	objTab_L.Screen.Send strFTP_pass & chr(13)
	objTab_L.Screen.WaitForString "ftp>"
	objTab_L.Screen.Send "binary" & chr(13)
	objTab_L.Screen.WaitForString "ftp>"
'---------------------------------------------
'   FTP TRANSFER main config L
'---------------------------------------------
	objTab_L.Screen.Send "cd " & " ./tested/" & strService & chr(13)
	If Not objTab_L.Screen.WaitForString ("is current directory", 60) Then
	    Call  TrDebug_No_Date ("FTP FOLDER " & "./tested/" & strService, "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
		exit sub		
	End If
	objTab_L.Screen.Send "put " & strConfigFileL & " " & strConfigFileL & chr(13)
	If Not objTab_L.Screen.WaitForString ("Successfully transferred", 60) Then
	    Call  TrDebug_No_Date ("L NODE FTP TRANSFER " & strConfigFileL, "FAILED", objDebug, MAX_LEN, 1, nInfo)
		exit sub		
    Else
	    Call  TrDebug_No_Date ("L NODE FTP TRANSFER " & strConfigFileL, "OK", objDebug, MAX_LEN, 1, nInfo)		
	End If
	objTab_L.Screen.WaitForString "ftp>"	
	'--------------------------------------------------------------------------------
    '  Start SSH session to Right Node
    '--------------------------------------------------------------------------------
	On Error Resume Next
	Err.Clear
	Set objTab_R = crt.Session.ConnectInTab("/S " & strFolder & strSessionR)
	If Err.Number <> 0 Then 
		Call  TrDebug_No_Date ("ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, MAX_LEN, 1, nInfo)
		Exit Sub
	End If
	On Error Goto 0
	objTab_R.Caption = strHostR
	objTab_R.Screen.Synchronous = True
	objTab_R.Screen.Send chr(13)
	nResult = objTab_R.Screen.WaitForString ("@" & strHostR & ">",2)
    If nResult = 0  Then
		If IsObject(objDebug) Then Call  TrDebug_No_Date (strHostR & " ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, nInfo) End If
		objTab_R.Session.Disconnect
    	Exit Sub
	End If
'---------------------------------------------
'   SAVE main config R
'---------------------------------------------
	objTab_R.Screen.Send "edit" & chr(13)
	objTab_R.Screen.WaitForString "@" & strHostR & "#"
	objTab_R.Screen.Send "save " & strConfigFileR & chr(13)
	objTab_R.Screen.WaitForString "@" & strHostR & "#"	
	objTab_R.Screen.Send "exit" & chr(13)
	objTab_R.Screen.WaitForString "@" & strHostR & ">"
	Call TrDebug_No_Date ("R NODE: SAVING CONFIGURATION LOCALLY: " & strConfigFileR ,"OK", objDebug, MAX_LEN, 1, nInfo)							
'---------------------------------------------
'   CONNECTING FTP R
'---------------------------------------------
	objTab_R.Screen.Send "ftp "  & strFTP_ip & chr(13)
	nResult = objTab_R.Screen.WaitForStrings (vWaitForFtp, 5)
     Select Case nResult
        Case 0
			Call  TrDebug_No_Date ("HOST " & strHostR, ": FTP CONNECTION TIME OUT", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("CONFIGURATION DOWNLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)					
'           objTab_L.Session.Disconnect			
'			crt.quit
			Exit Sub
        Case 1 
			Call  TrDebug_No_Date ("HOST " & strHostR & " HAS NO ROUTE TO FTP SERVER", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("1. Check FTP Server IP-address under MEF Loader Settings ", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("2. Make sure the Right node has a route to FTP server IP address ","", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("CONFIGURATION DOWNLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)		
	        objTab_R.Session.Disconnect			
    		crt.quit
    		Exit Sub
        Case 2 
			Call  TrDebug_No_Date ( "CONNECTING TO FTP", "OK", objDebug, MAX_LEN, 1, nInfo)   
		Case 3
			Call  TrDebug_No_Date ("FTP CONNECTION REFUSED", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("1. Make sure that FTP Server is running ", "", objDebug, MAX_LEN, 1, nInfo)
			Call  TrDebug_No_Date ("2. Make sure connection is not blocked by firewall","", objDebug, MAX_LEN, 1, nInfo)
			Call TrDebug_No_Date ("CONFIGURATION DOWNLOAD FAILED " , "", objDebug, MAX_LEN, 3, 1)		
	        objTab_R.Session.Disconnect			
    		crt.quit
    		Exit Sub	
     End Select		
	objTab_R.Screen.WaitForString "Name (" & strFTP_ip & ":" & strLogin & "):"
	objTab_R.Screen.Send strFTP_name & chr(13)
	objTab_R.Screen.WaitForString "Password:"
	objTab_R.Screen.Send strFTP_pass & chr(13)
	objTab_R.Screen.WaitForString "ftp>"
	objTab_R.Screen.Send "binary" & chr(13)
	objTab_R.Screen.WaitForString "ftp>"
'---------------------------------------------
'   FTP TRANSFER main config R
'---------------------------------------------
	objTab_R.Screen.Send "cd " & " ./tested/" & strService & chr(13)
	If Not objTab_R.Screen.WaitForString ("is current directory", 60) Then
	    Call  TrDebug_No_Date ("FTP FOLDER " & "./tested/" & strService, "NOT FOUND", objDebug, MAX_LEN, 1, nInfo)
		exit sub		
	End If
	objTab_R.Screen.Send "put " & strConfigFileR & " " & strConfigFileR & chr(13)
	If Not objTab_R.Screen.WaitForString ("Successfully transferred", 60) Then
	    Call  TrDebug_No_Date ("R NODE FTP TRANSFER " & strConfigFileR, "FAILED", objDebug, MAX_LEN, 1, nInfo)
		exit sub		
    Else
	    Call  TrDebug_No_Date ("R NODE FTP TRANSFER " & strConfigFileR, "OK", objDebug, MAX_LEN, 1, nInfo)  
	End If
	objTab_R.Screen.WaitForString "ftp>"
	'    crt.sleep 15000
    objTab_R.Session.Disconnect	
    objTab_L.Session.Disconnect	
	Call TrDebug_No_Date ("JOB DONE ", "", objDebug, MAX_LEN, 3, 1)	
	If IsObject(objDebug) Then objDebug.close : End If
	Set objFSO = Nothing
	Set objEnvar = Nothing
	crt.quit		
End Sub
'#######################################################################
' Function GetFileLineCount - Returns number of lines int the text file
'#######################################################################
 Function GetFileLineCount(strFileName, ByRef vFileLines, nDebug)
    Dim nIndex
	Dim strLine
	
    strFileWeekStream = ""	
	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		If 	InStr(strLine,"#") = 0 and InStr(strLine,"$") = 0 and strLine <> "" Then
			vFileLines(nIndex) = strLine
			If IsObject(objDebug) and nDebug = 1 Then objDebug.WriteLine "GetFileLineCount: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
			nIndex = nIndex + 1
		End If
	Loop
	objDataFileName.Close

    GetFileLineCount = nIndex
End Function
 '-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function

' ----------------------------------------------------------------------------------------------
'   Function  TrDebug_No_Date (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug_No_Date (strTitle, strString, objDebug, nChar, nFormat, nDebug)
Dim strLine
strLine = ""
If nDebug <> 1 Then Exit Function End If
If IsObject(objDebug) Then 
	Select Case nFormat
		Case 0
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) 
			strLine = strLine & ":  " & strTitle
			strLine = strLIne & strString
			objDebug.WriteLine strLine
			
		Case 1
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3)
			strLine = strLine & ":  " & strTitle
			If nChar - Len(strLine) - Len(strString) > 0 Then 
				strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
			Else 
				strLine = strLine & " " & strString
			End If
			objDebug.WriteLine strLine
		Case 2
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			
			If nChar - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
		Case 3
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine

	End Select
						
End If
End Function
'#######################################################################
 ' Function GetFileLineCountSelect - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCountSelect(strFileName, ByRef vFileLines,strChar1, strChar2, strChar3, nDebug)
    Dim nIndex
	Dim strLine, nCount, nSize
	Dim objDataFileName
	
    strFileWeekStream = ""	
	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	If nDebug = 1 Then objDebug.WriteLine "           GETTING SIZE OF THE FILE FIRST        "
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		Select Case Left(strLine,1)
			Case strChar1
			Case strChar2
			Case strChar3
			Case Else
					nIndex = nIndex + 1
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect:    " & strLine  End If  
		End Select
	Loop
	nCount = nIndex
	objDataFileName.close

    Redim vFileLines(nCount)
	nSize = UBound(vFileLines)
	Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	If nDebug = 1 Then objDebug.WriteLine "File Size: " & nCount & " Array Size: " & nSize
	If nDebug = 1 Then objDebug.WriteLine "           NOW TRYING TO RIGHT INTO AN ARRAY        "
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		Select Case Left(strLine,1)
			Case strChar1
			Case strChar2
			Case strChar3
			Case Else
					vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountSelect = nIndex
End Function
 '#######################################################################
 ' Function GetFileLineCountByGroup - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCountByGroup(strFileName, ByRef vFileLines, strGroup1, strGroup2, strGroup3, nDebug)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	
	GetFileLineCountByGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: -------------- GET FILE SIZE FIRST-----------" End If  
    Do While objDataFileName.AtEndOfStream <> True
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
'		strLine = All_Trim(objDataFileName.ReadLine)
		Select Case Left(strLine,1)
			Case "#"
'			Case "$"
			Case ""
			Case "["
				If strGroup1 = "All" Then 
					nGroupSelector = 1 
				Else 
					Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case "[" & strGroup2 & "]"
							nGroupSelector = 1
						Case "[" & strGroup3 & "]"
							nGroupSelector = 1
						Case Else
							nGroupSelector = 0
					End Select
				End If
			Case Else	
				If nGroupSelector = 1 Then nIndex = nIndex + 1 End If
		End Select
	Loop
	objDataFileName.Close
	Redim vFileLines(nIndex)
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: -------------- NOW READ TO ARRAY -----------" End If  
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
'		strLine = All_Trim(objDataFileName.ReadLine)
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
		Select Case Left(strLine,1)
			Case "#"
'			Case "$"
			Case ""
			Case "["
				If strGroup1 = "All" Then 
					nGroupSelector = 1 
				Else 
					Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case "[" & strGroup2 & "]"
							nGroupSelector = 1
						Case "[" & strGroup3 & "]"
							nGroupSelector = 1
						Case Else
							nGroupSelector = 0
					End Select
				End If
			Case Else	
				If nGroupSelector = 1 Then
					vFileLines(nIndex) = NormalizeStr(strLine, vDelim)
					' vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: vFileLines(" & nIndex & "): "  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
				End If
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountByGroup = nIndex
End Function
'-----------------------------------------------------------------
'     Function Normalize(strLine) - Removes all spaces around delimiters: arg1...arg3
'-----------------------------------------------------------------
Function NormalizeStr(strLine, vDelim)
Dim strNew
	strLine = LTrim(RTrim(strLine))
	strNew = ""
	For nInd = 0 to UBound(vDelim) - 1
		i = 0 
		Do While i <= UBound(Split(strLine,vDelim(nInd)))
				If UBound(Split(strLine,vDelim(nInd))) = 0 Then Exit Do End If
				If i < UBound(Split(strLine,vDelim(nInd))) Then strNew = strNew & LTrim(RTrim(Split(strLine,vDelim(nInd))(i))) & vDelim(nInd) End If
				If i = UBound(Split(strLine,vDelim(nInd))) Then strNew = strNew & LTrim(RTrim(Split(strLine,vDelim(nInd))(i))) End If
				i = i + 1
		Loop
		If i > 0 Then strLine = strNew End If
		strNew = ""
	Next
	NormalizeStr = strLine
End Function 
'-----------------------------------------------------------------
'     Function All_Trim(strLine) - Removes all speces form the string
'-----------------------------------------------------------------
Function All_Trim(strLine)
Dim nChar, strChar, i, strResult
	strResult = ""
	nChar = Len(strLine)
	For i = 1 to nChar
		strChar = Mid(strLine,i,1)
		If strChar <> " " Then strResult = strResult & strChar End If
	Next
		All_Trim = strResult
End Function
' ----------------------------------------------------------------------------------------------
'   Function  TrDebug_No_Date (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug_No_Date (strTitle, strString, objDebug, nChar, nFormat, nDebug)
Dim strLine
strLine = ""
If nDebug <> 1 Then Exit Function End If
If IsObject(objDebug) Then 
	Select Case nFormat
		Case 0
			strLine = ""
			strLine = strLine & ":  " & strTitle
			strLine = strLIne & strString
			objDebug.WriteLine strLine
			
		Case 1
			strLine = ""
			strLine = strLine & ":  " & strTitle
			If nChar - Len(strLine) - Len(strString) > 0 Then 
				strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
			Else 
				strLine = strLine & " " & strString
			End If
			objDebug.WriteLine strLine
		Case 2
			strLine = ""
			
			If nChar - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
		Case 3
			strLine = ""
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine
			strLine = ""
			If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
			strLine = ""
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine

	End Select
						
End If
End Function
'----------------------------------------------------------------
'   Function FocusToParentWindow(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function FocusToParentWindow(strPID)
Dim objShell
Call TrDebug_No_Date ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 0) 
Const IE_PAUSE = 70
	Set objShell = CreateObject("WScript.Shell")
	crt.sleep IE_PAUSE  
	objShell.SendKeys "%"	
	crt.sleep IE_PAUSE
	objShell.AppActivate strPID			
	crt.sleep IE_PAUSE  
	objShell.SendKeys "% "
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function GetAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetAppPID(ByRef strPID, ByRef strParentPID, strAppName)
Dim objWMI, colItems
Const IE_PAUSE = 70
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug_No_Date ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 0)
				On error Goto 0 
				Exit Do
		End If 
'		wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'Launcher Ver.'"  WHERE Name = 'iexplore.exe' OR Name = 'wscript.exe'
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug_No_Date ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug_No_Date ("GetMyPID: RESTORE IE WINDOW:", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 0) 
			If pUser = strUser then 
				strPID = process.ProcessId
				strParentPID = Process.ParentProcessId
'				Call TrDebug_No_Date ("GetMyPID: ", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
'				Call TrDebug_No_Date ("GetMyPID: ", "Caption: " & process.Caption & ", CSName " & process.CSName & ", Description: " & process.Description & ", Handle: " &  Process.Handle,objDebug, MAX_LEN, 1, 1) 
			GetAppPID = True
				Exit For
			End If
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUserSYS
'----------------------------------------------------------------------------------
Function GetScreenUserSYS()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserSYS = strScreenUser
End Function
'----------------------------------------------
'Function  GetValue(nInd)
'----------------------------------------------
Function GetValue(nInd)
Dim i
	GetValue = UNKNOWN
	If IsNumeric(nInd) Then 
		GetValue = Trim(Split(vSettings(nInd),"=")(1))
	Else 
		For i=0 to UBound(vParamNames) - 1
		    If vParamNames(i) = nInd Then 
				GetValue = Trim(Split(vSettings(i),"=")(1))
				Exit For
			End If 
		Next 
	End If
End Function
'---------------------------------------------------------
'  Function LoadSettings(strFileSettings, strFileParam)
'---------------------------------------------------------
Function LoadSettings(strFileSettings, ByRef strFileParam)
Dim nSettings, vLines, nInd
	If objFSO.FileExists(strFileSettings) Then 
		nSettings = GetFileLineCountByGroup(strFileSettings, vLines,"Settings","Sessions","Templates",0)
		For nInd = 0 to nSettings - 1 
			For i = 0 to UBound(vSettings) - 1
				Select Case i
					Case 5,6,15,16
					Case else
						If vParamNames(i) = Trim(Split(vLines(nInd),"=")(0)) Then 
							vSettings(i) = vLines(nInd)
						End If
				End Select 
			Next
		Next
	Else 
	    LoadSettings = -2
		Exit Function
	End If 
	strFileParam = GetValue(CONFIGS_PARAM)
	If objFSO.FileExists(strFileParam) Then 
		nSettings = GetFileLineCountByGroup(strFileParam, vLines,"Settings","Sessions","Templates",0)
		For nInd = 0 to nSettings - 1 
			For i = 0 to UBound(vSettings) - 1
				Select Case i
					Case 5,6,15,16
					Case else
						If vParamNames(i) = Trim(Split(vLines(nInd),"=")(0)) Then 
							vSettings(i) = vLines(nInd)
						End If
				End Select 
			Next
		Next
	Else 
	    LoadSettings = -1
		Exit Function
	End If
    Call SetValue(5,strCRT_InstallFolder,False)
	Call SetValue(6,strDirectoryWork,False)
	Call SetValue(15,strCRT_SessionFolder,False)
	Call SetValue(16,strFTP_Folder,False)
	strDirectoryConfig = GetValue(7)
	If GetValue(17) = "True" Then nWindowState = HideTerminal Else nWindowState = ShowTerminal
	LoadSettings = 1
End Function 
'----------------------------------------------
'Function  SetValue(nInd)
'----------------------------------------------
Function SetValue(nInd,strValue,bNormal)
Dim i, nTab
    If bNormal Then nTab = "" Else nTab = Space(30 - Len(vParamNames(nInd)))
	SetValue = False
	If IsNumeric(nInd) Then 
		vSettings(nInd) = vParamNames(nInd) & Space(30 - Len(vParamNames(nInd))) & "= " & strValue
		SetValue = True
	Else 
		For i=0 to UBound(vParamNames) - 1
		    If vParamNames(i) = nInd Then 
				vSettings(i) = vParamNames(i) & Space(30 - Len(vParamNames(i))) & "= " & strValue
				SetValue = True
				Exit For
			End If 
		Next 
	End If
End Function
'--------------------------------------------
'  Sub SetOriginalSettings(vSettings)
'--------------------------------------------
Sub SetOriginalSettings(ByRef vSettings)
Dim nInd
	For nInd = 0 to UBound(vParamNames) - 1
		Call SetValue(nInd,"Unknown",False)
	Next
End Sub 
