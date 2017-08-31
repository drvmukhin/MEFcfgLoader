'----------------------------------------------------------------------------------
'            JUNIPER MEF CONFIGURATION LOADER
'----------------------------------------------------------------------------------

Const LDR_SCRIPT_NAME = "MEF_Loader_"
Const ForAppending = 8
Const ForWriting = 2
Const HttpTextColor1 = "#292626"
Const HttpTextColor2 = "#F0BC1F"
Const HttpTextColor3 = "#EBEAF7"
Const HttpTextColor4 = "#A4A4A4"
Const HttpBgColor1 = "Grey"
Const HttpBgColor2 = "#272727" 
Const HttpBgColor3 = "#2C2A23" 
Const HttpBgColor4 = "#504E4E"
'Const HttpBgColor5 = "#0D057F"
Const HttpBgColor5 = "#101010"
Const HttpBgColor6 = "#8B9091"
Const HttpBdColor1 = "Grey"
Const PAUSE_TIME = 		20
Const MAX_TIMER = 		20
Const MAX_TIME_DELTA = 	180
Const MAX_LEN = 		134
Const nABC =			2
Const HideTerminal = 7
Const ShowTerminal = 1
'
'   Global Array Variables
Dim vConfigFileLeft,vFilterList, vPolicerList, vCIR, vCBS, vFilterProfileList
Dim FltCurrent1, FltCurrent2
Dim CIRo, PIRo, CBSo, CBSdef, PBSdef, PBSo, MTU
'
'   Global Variables
Dim D0 
Dim nResult, nTail, nIndex, nCountWeek, nCountMonth
Dim strLine
'
'   Global FileSystem Variables
Dim strFileSessionTmp, strFileSession, strCRTexe
Dim strDirectoryWork,strDirectoryTmp, strDirectoryLog,strDirectoryBin,strDirectoryBackUp, strDirectoryConfig, strCRT_InstallFolder 
Dim strFTP_Folder, SecureCRT, strDirectoryScript, strCRT_SessionFolder, strDirectoryTested
'
'   Other Global Variables
Dim	strServiceID, strFlavorID, strTaskID
Dim nWindowState
Dim nLine
Dim intX, intY
Dim nSession, nSessionTmp, nInventory
Dim vSession, vSessionTmp(20), vMsg(20)
Dim vLine
Dim vvMsg(20,3)
Dim strFolder
Dim objFSO, objEnvar, objDebug, objShell
Dim vIE_Scale
Dim UserConfigFile
Dim objShellApp
Dim strPID
Dim IE_Window1_JustCreated
Dim IE_Reload_Form1
Redim IE_Reload_Form1(10)
'----------------------------------
' NEW VARIABLES
'----------------------------------
Dim vFlavors, vSvc, vNodes(2,5), vTemplates(4), vSettings, vOld_Settings
Dim strConfigFileL, strConfigFileR, strVersion
Dim Platform, DUT_Platform
Dim VBScript_DNLD_Config, VBScript_Upload_Config, VBScript_FWF_Apply, VBScript_FTP_User, VBScript_Set_Node, VBScript_BWP_Apply, VBScript_FLT_and_BWP_Apply
Dim strTempOrigFolder, strTempDestFolder, vXLSheetPrefix(4)
Dim objXLS, objWrkBk, objXLSeet
Dim vDelim, vParamNames,vPlatforms, objWMIService, IE_Window_Title
Dim objFolder, colFiles, IPConfigSet, strEditor, SecureCRT_Installed, FileZilla_Installed
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
	Const MAIN_TITLE = "Juniper Networks MEF Configuration Loader"
    Const CRT_REG_INSTALL = "HKEY_LOCAL_MACHINE\SOFTWARE\VanDyke\SecureCRT\Install\Main Directory"
    Const CRT_REG_SESSION = "HKEY_CURRENT_USER\Software\VanDyke\SecureCRT\Config Path"
	Const MEF_CFG_LDR_REG = "HKEY_CURRENT_USER\Software\JnprMCL\Main Directory"
	Const FTP_REG_INSTALL = "HKEY_CURRENT_USER\Software\FileZilla Server\Install_Dir"
	Const NOTEPAD_PP = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe\"
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"

vDelim = Array("=",",",":")
ReDim vSession (10)
Redim vSettings(30)
ReDim vOld_Settings(UBound(vSettings))
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
	
strFileSessionTmp = "<WorkDir>\Temp\sessions_tmp.dat"
strDirectoryWork = "<WorkDir>"
strDirectoryConfig = "<ConfigDir>"
strDirectoryBackUp = "<ConfigDir>\backup"
strDirectoryTested = "<ConfigDir>\tested"
strDirectoryScript = "<WorkDir>\script"
strDirectoryLog = "<WorkDir>\log"
strDirectoryTmp = "<WorkDir>\Temp"
strDirectoryBin = "<WorkDir>\Bin"
strFileSession = "<WorkDir>\config\session.ini"
FILE_DEBUG = "Log\debug-mefloader-<date>.log"
strFileSettings = "<WorkDir>\config\settings_new.dat"
strFileParam = "<WorkDir>\config\topology.dat"
strFTP_Folder = "C:"
strCRT_SessionFolder = "C:"
strCRT_InstallFolder = "C:\Program Files\VanDyke Software\SecureCRT"
SecureCRT = "SecureCRT.exe"
strCRTexe = "\SecureCRT.exe"""


strConfigFileL = ""
strConfigFileR = ""
Platform = "acx"
DUT_Platform = "Unknown"
D0 = DateSerial(2015,1,1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = WScript.CreateObject("WScript.Shell")
Set objShell = WScript.CreateObject("WScript.Shell")
'
' DEBUG AND LOGS VARIABLES
Dim nDebug, nInfo, nDebug_, ShowLog, nDebugCRT, nDay, bVerbose
nDebug = 0
nDebug_ = 1
nInfo = 1
nDay = Day(Date())
ShowLog = False
bVerbose = False
nDebugCRT = 0 

CurrentDate = Date()
CurrentTime = Time()
D0 = DateSerial(2015,1,1)

Main()
  
If IsObject(objDebug) Then objDebug.Close : End If
Set objFSO = Nothing
set objEnvar = Nothing
Set objShell = Nothing

Sub Main()
	'
	'  GET SCREEN RESOLUTION
	Call GetScreenResolution(vIE_Scale, 0)
	'
	'  Set work directory
	strDirectoryWork =  objFSO.GetParentFolderName(Wscript.ScriptFullName)
	'  Define file system variables
	FILE_DEBUG         = Replace(FILE_DEBUG,"<WorkDir>",strDirectoryWork)
	strDirectoryLog    = Replace(strDirectoryLog,"<WorkDir>",strDirectoryWork)
	strDirectoryTmp    = Replace(strDirectoryTmp,"<WorkDir>",strDirectoryWork)
	strDirectoryBin    = Replace(strDirectoryBin,"<WorkDir>",strDirectoryWork)
	strDirectoryScript = Replace(strDirectoryScript,"<WorkDir>",strDirectoryWork)
	strFileSession     = Replace(strFileSession,"<WorkDir>",strDirectoryWork)
	strFileSettings    = Replace(strFileSettings,"<WorkDir>",strDirectoryWork)
	strFileParam       = Replace(strFileParam,"<WorkDir>",strDirectoryWork)
	strFileSessionTmp  = Replace(strFileSessionTmp,"<WorkDir>",strDirectoryWork)
	strDebugFile = Replace(FILE_DEBUG,"<date>",Year(Date()) & Month(Date) & Day(Date()))
	If Not ValidateFS(nDebug) then Exit Sub
	'
	'  Start Log Session
	If Not OpenLogSession(objDebug, strDebugFile, strDirectoryBin, False, ShowLog, bVerbose) Then Exit Sub
	'
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
		On Error Resume Next
			Err.Clear
			IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
			if Err.Number <> 0 Then 
				IE_Window_Title = "Internet Explorer"
			End If
		On Error Goto 0
	'
	'  CHOOSE A DEFAULT TEXT EDITOR
	On Error Resume Next
		Err.Clear
		strEditor =  """" & objShell.RegRead(NOTEPAD_PP) & """"
		if Err.Number <> 0 Then 
			strEditor = "notepad.exe"
		End If
	On Error Goto 0
	'
	'  	Validate FTP Server (FileZilla) is installed on the system
	FileZilla_Installed = VaildateFTP(strFTP_Folder)
	'
	'  Validate SecureCRT Installation on the local system
	SecureCRT_Installed = ValidateSecureCRT(strCRT_SessionFolder,strCRT_InstallFolder)
	strCRTexe = """" & strCRT_InstallFolder & strCRTexe
	'   Set original values for the settings values
	Call SetOriginalSettings(vSettings)	
	'   Load settings and topology 
    Select Case LoadSettings(strFileSettings, strFileParam)
	    Case -1
			MsgBox "Cant' find Settings file: " & chr(13) & strFileParam
			Exit Sub
		Case -2 
			MsgBox "Cant' find Settings file: " & chr(13) & strFileSettings
			Exit Sub
	End Select
	'   Get names of telnet scripts
	Call LoadScriptsNames()
    '   Get Version of the CfgLoader
	Call GetCfgLoaderVersion()
	'   Load settings for XLS Templates
	Call GetXLStemplates()
	'   Get List of Supported platforms
	nPlatform = GetFileLineCountByGroup(strFileSettings, vPlatforms,"Supported_Platforms","","",0)
	'  	Load Test Bed Topology
    Call LoadTopology(strFileParam,vNodes)
	' Load test series 
	nService = GetTestSeries(strFileParam, vSvc, vFlavors, nDebug)
	Call GetTestCases(vSvc, vFlavors, 1)
    ' Create folder structure for backup and tested configuration files
	Call CreateFoldersForTestedConfigurations(strDirectoryConfig)
	' Load data for last tested series
	Call LoadLastTestedSeries(strFileSessionTmp)
	'----------------------------
	'   Open main form
	'----------------------------
	'
	'  Create IExplorer Window for Main Form
	IE_Window1_JustCreated = Set_IE_obj(main_objIE)
	Call FormReload(main_objIE, IE_Reload_Form1, False)
	
	Do
        nResult = IE_PromptForInput(main_objIE, IE_Window1_JustCreated, IE_Reload_Form1,vIE_Scale, nDebug)
		Select Case nResult
			Case 0,-1
                If IsObject(main_objIE) Then 
					Set main_objIE = Nothing 
				End If
				Exit Sub
			Case 2
				Call FormReload(main_objIE, IE_Reload_Form1, True)
				Select Case LoadSettings(strFileSettings, strFileParam)
					Case -1
						MsgBox "Cant' find Settings file: " & chr(13) & strFileParam
						Exit Sub
					Case -2 
						MsgBox "Cant' find Settings file: " & chr(13) & strFileSettings
						Exit Sub
				End Select
				'   Get names of telnet scripts
				Call LoadScriptsNames()
				'   Get Version of the CfgLoader
				Call GetCfgLoaderVersion()
				'   Load settings for XLS Templates
				Call GetXLStemplates()
				'   Get List of Supported platforms
				nPlatform = GetFileLineCountByGroup(strFileSettings, vPlatforms,"Supported_Platforms","","",0)
				'  	Load Test Bed Topology
				Call LoadTopology(strFileParam,vNodes)
				' Load test series 
				nService = GetTestSeries(strFileParam, vSvc, vFlavors, nDebug)
				Call GetTestCases(vSvc, vFlavors, 0)
				' Create folder structure for backup and tested configuration files
				Call CreateFoldersForTestedConfigurations(strDirectoryConfig)
				' Load data for last tested series
				Call LoadLastTestedSeries(strFileSessionTmp)			
			Case Else 
        End Select
	Loop
	objDebug.Close
End Sub
'-----------------------------------------------------------------------------
'      Function Displays a Message with OK Button. Returns True.
'-----------------------------------------------------------------------------
 Function IE_MSG (vIE_Scale, strTitle, ByRef vLine, ByVal nLine, objIEParent)
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim nDebug, cellW, CellH
	Dim g_objIE, objShell
    Set g_objIE = Nothing
    Set objShell = Nothing
	nDebug = 0
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_MSG = True
	Call IE_Hide(objIEParent)
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nRatioX = intX/1920
    nRatioY = intY/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((130 + nLine * 35) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	
 '  If nDebug = 1 Then MsgBox "intX=" & intX & "   intY=" & intY & "   RatioX=" & nRatioX & "  RatioY=" & nRatioY & "   Cell Width=" & cellW & "  Cell Hight=" & cellH End If

	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	strHTMLBody = "<br>"
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center;font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
	Next		
	
    strHTMLBody = strHTMLBody &_
                "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
				"; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; left: " & Int((CellW - nButtonX)/2) & "px; bottom: 4px' name='OK' AccessKey='O' onclick=document.all('ButtonHandler').value='OK';><u>O</u>K</button>" & _
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"

			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.Visible = True
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title

	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strMyPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then strMyPID = Split(strLine,""",""")(1)
	     ' Call TrDebug("READ TASK PID:" , strLine, objDebug, MAX_LEN, 1, 1)
    Next
    If strMyPID = "" Then Call GetAppPID(strMyPID, "iexplore.exe")
	objShell.AppActivate strMyPID
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "OK"
				IE_MSG = True
				g_objIE.quit
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
		Call IE_UnHide(objIEParent)
End Function
'--------------------------------------------------------------------------------
'   Function IE_MSG_Internal(ByRef g_objIE, vIE_Scale, strTitle, vLine, ByVal nLine)
'      displays message window inside the parent form
'--------------------------------------------------------------------------------
 Function IE_MSG_Internal (vIE_Scale, strTitle, vLine, ByVal nLine,ByRef g_objIE)
    Dim strHTMLBody
	Dim intX
    Dim intY
	Dim WindowH, WindowW, CellH, CellW, IE_Window_Title, nTab, nRatioX, nRatioY, BottomH
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim strPID
	Dim nButtonX, nButtonY
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	IE_Window_Title = strTitle
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	IE_MSG_Internal = True
	LoginTitleH = 40
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW
	WindowH = CellH
	LineH = 24
	
	nTab = Round(20 * nRatioX,0)
	BottomH = Round(10 * nRatioY,0)
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If

	Dim BackGroundColor,ButtonColor ,TitleBgColor,TitleTextColor,ParentWindowH,ParentWindowW
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	TitleBgColor = HttpBgColor2
	TitleTextColor = HttpTextColor2
	MainTextColor = HttpTextColor1
    ParentWindowH = g_objIE.height
	ParentWindowW =	g_objIE.width
	'
	'  Draw a message board content
    strHTMLBody = strHTMLBody &_
		"<table valign=""middle"" border=""0"" cellpadding=""1"" cellspacing=""1"" style="" position: relative; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border-width: 1px; border-color: transparent; background-color: Transparent;" &_
		"width: " & CellW & "px; height: " & CellH & "px"">" & _
			"<tbody>" & _
				"<tr height=""" & LoginTitleH & """ >" &_
					"<td align=""center"" style=""background-color: " & TitleBgColor & "; color: " & TitleTextColor & ";"">" & IE_Window_Title & "</td>" &_
				"</tr>" &_
				"<tr height=""" & LineH & """ ></tr>"
				For nInd = 0 to nLine - 1
					If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
					strHTMLBody = strHTMLBody &_
					"<tr height=""" & LineH & """ >" &_
						"<td>" &_ 
							"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" &_
								vLine(nInd,0) &_
							"</p>" &_
						"</td>" &_
					"</tr>" 					
				Next		
	strHTMLBody = strHTMLBody &_				
				"<tr height=""" & LineH & """ ></tr>" &_				
				"<tr><td align=""center"">" &_
					"<button class=""ok_button"" style=""" &_ 
					"width:" & nButtonX & "px; height:" & nButtonY & "px;""" &_
					"id='MY_OK' name='MY_OK' AccessKey='O' onclick=document.all('My_ButtonHandler').value='OK';><u>O</u>K</button>" & _
					"<input name='My_ButtonHandler' type='hidden' value='Nothing Clicked Yet'></td>" &_
				"</tr>" &_ 
			"</tbody></table>"	

	'
	'  Load Message board form
	g_objIE.Document.getElementById("divMessage").innerHTML  = strHTMLBody
	'
	'  Shadow Div Properties
	Call IE_Set_Window_Shadow(g_objIE,False)	
	'
	'  Message Board Properties
	'g_objIE.Document.getElementById("divMessage").style.opacity  = 0.95
	g_objIE.Document.getElementById("divMessage").style.position  = "absolute"
	g_objIE.Document.getElementById("divMessage").style.top = (ParentWindowH - CellH)/2 & "px"
	g_objIE.Document.getElementById("divMessage").style.left = (ParentWindowW - CellW)/2 & "px"
	g_objIE.Document.getElementById("divMessage").style.backgroundColor = "grey"
	g_objIE.Document.getElementById("divMessage").style.color  = HttpTextColor2	
	g_objIE.Document.getElementById("divMessage").style.fontsize  = "20px"
	g_objIE.Document.getElementById("divMessage").style.fontfamily = "arial narrow"	
	g_objIE.Document.getElementById("divMessage").style.height  = CellH & "px"	
	g_objIE.Document.getElementById("divMessage").style.width  = CellW & "px"
	g_objIE.Document.getElementById("divMessage").style.boxShadow="10px 10px 10px #101010"
	'
	'  Main Cycle of the Message Form 
	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("My_ButtonHandler").Value
		if Err.Number <> 0 then 
			MsgBox Err.Description
			IE_MSG_Internal = False 
			exit do
		End If 
		If g_objIE.width <> ParentWindowW Then g_objIE.width = ParentWindowW End If
		If g_objIE.height <> ParentWindowH Then g_objIE.height = ParentWindowH End If		
		On Error Goto 0
		Select Case strNothing
			Case "OK"
				IE_MSG_Internal = True
				Exit Do
		End Select
		Wscript.Sleep 200
	Loop
	'
	'  Close Shadow Div and Message Board
	Call IE_Clear_Window_Shadow(g_objIE)	
End Function
'-----------------------------------------------------------------------------------------
'      Function IE_CONT_Internal
'-----------------------------------------------------------------------------------------
 Function IE_CONT_Internal(vIE_Scale, strTitle, vLine, ByVal nLine, ByRef g_objIE,nDebug)
    Dim strHTMLBody
	Dim intX
    Dim intY
	Dim WindowH, WindowW, CellH, CellW, IE_Window_Title, nTab, nRatioX, nRatioY, BottomH
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim strPID
	Dim nButtonX, nButtonY
	IE_CONT_Internal = False
	'
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	IE_Window_Title = strTitle
	'
	' SCREEN RESOLUTION
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW
	WindowH = CellH
	LineH = 24
	LoginTitleH = 40
	nTab = Round(20 * nRatioX,0)
	BottomH = Round(10 * nRatioY,0)
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
    '
    '   MAIN COLORS OF THE FORM
	Dim BackGroundColor,ButtonColor ,TitleBgColor,TitleTextColor,ParentWindowH,ParentWindowW
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	MainTextColor = HttpTextColor1
	TitleBgColor = HttpBgColor2
	TitleTextColor = HttpTextColor2
    ParentWindowH = g_objIE.height
	ParentWindowW =	g_objIE.width
	'
	'  Draw a message board content
    strHTMLBody = strHTMLBody &_
		"<table valign=""middle"" border=""0"" cellpadding=""1"" cellspacing=""1"" style="" position: relative; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border-width: 1px; border-color: transparent; background-color: Transparent;" &_
		"width: " & CellW & "px; height: " & CellH & "px"">" & _
			"<tbody>" & _
				"<tr height=""" & LoginTitleH & """ >" &_
					"<td colspan=""2"" align=""center"" style=""background-color: " & TitleBgColor & "; color: " & TitleTextColor & ";"">" & IE_Window_Title & "</td>" &_
				"</tr>" &_
				"<tr height=""" & LineH & """ ></tr>"
				For nInd = 0 to nLine - 1
					If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
					strHTMLBody = strHTMLBody &_
					"<tr height=""" & LineH & """ >" &_
						"<td colspan=""2"" >" &_ 
							"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" &_
								vLine(nInd,0) &_
							"</p>" &_
						"</td>" &_
					"</tr>" 					
				Next		
    strHTMLBody = strHTMLBody &_
				"<tr height=""" & LineH & """ ></tr>" &_
				"<tr>" &_
					"<td align=""center"">" &_				
						"<button class=""ok_button"" style='" &_
						"width:" & nButtonX & ";height:" & nButtonY & ";" &_
						"position: absolute; left: " & nTab & "px; bottom: " & BottomH & "px'" &_
						"id='MY_OK' name='MY_OK' AccessKey='Y' onclick=document.all('My_ButtonHandler').value='YES';><u>Y</u>ES</button>" & _
					"</td>" &_
					"<td align=""center"">" &_				
						"<button class=""ok_button"" style='" &_
						"width:" & nButtonX & ";height:" & nButtonY & ";" &_
						"position: absolute; right: " & nTab & "px; bottom: " & BottomH & "px'" &_
						"id='MY_NO' name='MY_NO' AccessKey='N' onclick=document.all('My_ButtonHandler').value='NO';><u>N</u>O</button>" & _
					"</td>" &_
					"<input name='My_ButtonHandler' type='hidden' value='Nothing Clicked Yet'>" &_
				"</tr>" &_
			"</tbody></table>"					
	'
	'  Load Message board form
	g_objIE.Document.getElementById("divMessage").innerHTML  = strHTMLBody
	'
	'  Shadow Div Properties
	Call IE_Set_Window_Shadow(g_objIE,False)	
	'
	'  Message Board Properties
	'g_objIE.Document.getElementById("divMessage").style.opacity  = 0.95
	g_objIE.Document.getElementById("divMessage").style.position  = "absolute"
	g_objIE.Document.getElementById("divMessage").style.top = (ParentWindowH - CellH)/2 & "px"
	g_objIE.Document.getElementById("divMessage").style.left = (ParentWindowW - CellW)/2 & "px"
	g_objIE.Document.getElementById("divMessage").style.backgroundColor = "grey"
	g_objIE.Document.getElementById("divMessage").style.color  = HttpTextColor2	
	g_objIE.Document.getElementById("divMessage").style.fontsize  = "20px"
	g_objIE.Document.getElementById("divMessage").style.fontfamily = "arial narrow"	
	g_objIE.Document.getElementById("divMessage").style.height  = CellH & "px"	
	g_objIE.Document.getElementById("divMessage").style.width  = CellW & "px"
	g_objIE.Document.getElementById("divMessage").style.boxShadow="10px 10px 10px #101010"
	'
	'  Main Cycle of the Message Form 
	Do
		On Error Resume Next
		If g_objIE.width <> ParentWindowW Then g_objIE.width = ParentWindowW End If
		If g_objIE.height <> ParentWindowH Then g_objIE.height = ParentWindowH End If						
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("My_ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("My_ButtonHandler").Value
			Case "NO"
				IE_CONT_Internal = False
				Exit Do
			Case "YES"
				IE_CONT_Internal = True
				Exit Do
		End Select
		Wscript.Sleep 200
	Loop
	'
	'  Close Shadow Div and Message Board
	Call IE_Clear_Window_Shadow(g_objIE)
End Function
'--------------------------------------------------------------------------------
'   Function IE_Set_Window_Shadow(ByRef g_objIE)
'--------------------------------------------------------------------------------
 Function IE_Set_Window_Shadow(ByRef g_objIE, bWait)
	Dim ParentWindowH,ParentWindowW, Hidden, wait_left, wait_top
    ParentWindowH = g_objIE.height
	ParentWindowW =	g_objIE.width
	
	Hidden = True
	If bWait Then Hidden = False
	'
	'  Shadow Div Properties
	g_objIE.Document.getElementById("divShadow").style.opacity = 0.6
	g_objIE.Document.getElementById("divShadow").style.position  = "absolute"
	g_objIE.Document.getElementById("divShadow").style.top = "0px"
	g_objIE.Document.getElementById("divShadow").style.left = "0px"
	g_objIE.Document.getElementById("divShadow").style.backgroundColor = "black"
	g_objIE.Document.getElementById("divShadow").style.fontsize  = "20px"
	g_objIE.Document.getElementById("divShadow").style.fontfamily  = "arial narrow"	
	g_objIE.Document.getElementById("divShadow").style.height  = ParentWindowH & "px"	
	g_objIE.Document.getElementById("divShadow").style.width  = ParentWindowW & "px"
	g_objIE.Document.getElementById("divTransparent").style.position  = "absolute"
	g_objIE.Document.getElementById("divTransparent").style.top = "0px"
	g_objIE.Document.getElementById("divTransparent").style.left = "0px"
	g_objIE.Document.getElementById("divTransparent").style.backgroundColor = "transparent"
	g_objIE.Document.getElementById("divTransparent").style.height  = ParentWindowH & "px"	
	g_objIE.Document.getElementById("divTransparent").style.width  = ParentWindowW & "px"
	wait_left = (ParentWindowW - 100)/2 
	wait_top  = (ParentWindowH - 100)/2 	
	g_objIE.Document.getElementById("wait_icon").hidden  = Hidden
	g_objIE.Document.getElementById("wait_icon").style.position  = "absolute"
	g_objIE.Document.getElementById("wait_icon").style.left = wait_left & "px"	
	g_objIE.Document.getElementById("wait_icon").style.top  = wait_top & "px"

End Function
'--------------------------------------------------------------------------------
'   Function IE_Clear_Window_Shadow(ByRef g_objIE)
'--------------------------------------------------------------------------------
 Function IE_Clear_Window_Shadow(ByRef g_objIE)
	Dim ParentWindowH,ParentWindowW
    ParentWindowH = g_objIE.height
	ParentWindowW =	g_objIE.width
	'
	'  Clear Shadow
	g_objIE.Document.getElementById("divShadow").style.left = 1 * (ParentWindowW + nTab) & "px"
	g_objIE.Document.getElementById("divShadow").style.top = "0px"	
	g_objIE.Document.getElementById("divShadow").style.zindex = "1"
	'
	'  Clear Message Board
	g_objIE.Document.getElementById("divMessage").innerHTML  = ""	
	g_objIE.Document.getElementById("divMessage").style.backgroundColor = "transparent"	
	g_objIE.Document.getElementById("divMessage").style.left = 1 * (ParentWindowW + nTab) & "px"
	g_objIE.Document.getElementById("divMessage").style.top = "0px"
	g_objIE.Document.getElementById("divMessage").style.zindex = "3"
	'
	'  Clear Transparent Div
	g_objIE.Document.getElementById("divTransparent").style.left = 1 * (ParentWindowW + nTab) & "px"
	g_objIE.Document.getElementById("divTransparent").style.top = "0px"	
	g_objIE.Document.getElementById("divTransparent").style.zindex = "2"
	'g_objIE.Document.getElementById("divTransparent").innerHTML  = ""
	'
	'  Hide wait icon
	g_objIE.document.getElementById("wait_icon").hidden  = True	
End Function
'###################################################################################
' Function returns True if object/string exists in data file                 
'###################################################################################
Function MyObjectExist( byRef strFilePath, byRef strObjectName)
	    MyObjectExist = False
		Set objFileObject = objFSO.OpenTextFile(strFilePath)
		Do While objFileObject.AtEndOfStream <> True
            vLine = Split(objFileObject.ReadLine,",") 
	        If vLine(0) = strObjectName Then 
                MyObjectExist = True
			End If
        Loop
	    objFileObject.Close 
End Function
'###################################################################################
'  Authenticate User against its password. Requires an account data file as input                  
'###################################################################################
Function Authenticate( byRef strFilePath, byRef strObjectName, byRef passwd)
	    Authenticate = False
		Set objFileObject = objFSO.OpenTextFile(strFilePath)
		Do While objFileObject.AtEndOfStream <> True
            vLine = Split(objFileObject.ReadLine,",") 
	        If vLine(0) = strObjectName Then 
                If passwd = vLine(2) Then Authenticate = True End If
			End If
        Loop
	    objFileObject.Close 
End Function
'###################################################################################
'  Function MinQ - Returs the Minimum of two numeric values                  
'###################################################################################
Function MinQ( nA, nB)
   If nA < nB Then 
     MinQ = nA 
   Else 
     MinQ = nB
   End If
End Function
'###################################################################################
'  Function MinQ - Returs the Minimum of two numeric values                  
'###################################################################################
Function MaxQ( nA, nB) 
   If nA > nB Then 
     MaxQ = nA 
   Else 
     MaxQ = nB
   End If
End Function
'#######################################################################
 ' Function GetFileLineCount - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCount(strFileName, ByRef vFileLines, nDebug)
    Dim nIndex
	Dim strLine
	Dim objDataFileName
	
    strFileWeekStream = ""	
	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		If 	Left(strLine,1)<>"#" Then
			vFileLines(nIndex) = strLine
			If nDebug = 1 Then objDebug.WriteLine "GetFileLineCount: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
			nIndex = nIndex + 1
		End If
	Loop
	objDataFileName.Close
    GetFileLineCount = nIndex
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
'-------------------------------------------------------------
'    Function GetScreenResolution(vIE_Scale, intX,intY)
'-------------------------------------------------------------
Function GetScreenResolution(ByRef vIE_Scale, nDebug)
Dim g_objIE, f_objShell, intX, intY, intXreal, intYreal	
Redim vIE_Scale(2,3)
	nInd = 0
	Call Set_IE_obj (g_objIE)
	
	With g_objIE
		.Visible = False
		.Offline = True	
		.navigate "about:blank"
		Do
			WScript.Sleep 200
		Loop While g_objIE.Busy	
		.Document.Body.innerHTML = "<p>tEST</p>"
		.MenuBar = False
		.StatusBar = False
		.AddressBar = False
		.Toolbar = False		
		.Document.body.scroll = "no"
		.Document.body.Style.overflow = "hidden"
		.Document.body.Style.border = "None " & HttpBdColor1
		.Height = 100
		.Width = 100
		OffsetX = .Width - .Document.body.clientWidth
		OffsetY = .Height - .Document.body.clientHeight
		.FullScreen = True
		.navigate "about:blank"	
		 intXreal = .width
		 intYreal = .height
		.Quit
	End With
	If intXreal => 1440 Then intX = 1920 else intX = intXreal
	If intYreal => 900 Then intY = 1080  else intY = intYreal
	vIE_Scale(0,0) = intX : vIE_Scale(0,1) = OffsetX : vIE_Scale(0,2) = intXreal 
	vIE_Scale(1,0) = intY : vIE_Scale(1,1) = OffsetY : vIE_Scale(1,2) = intYreal
	
	Set g_objIE = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------
' Function returns the number of the line from 1 to N which contains string strObject. Returns 0 if nothing found
'------------------------------------------------------------------------------------------------------------------
Function GetObjectLineNumber( byRef vArray, nArrayLen, byRef strObjectName)
Dim nInd
	nInd = 0
	GetObjectLineNumber = 0
	Do While nInd < nArrayLen
	If InStr(vArray(nInd), strObjectName) <> 0	Then 
		GetObjectLineNumber = nInd + 1
		Exit Do
	End If
	nInd = nInd + 1
    Loop
End Function
' ----------------------------------------------------------------------------------------------
'   Function  TrDebug (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug (strTitle, strString, objDebug, nChar, nFormat, nDebug)
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
'-------------------------------------------------
'  Function NormalizePatterm(strPattern)
'-------------------------------------------------
Function NormalizePatterm(ByRef strPattern, EmptyString)
Dim vSpecualChars, Special
	vSpecualChars = Array("\","$",".","[","]","{","}","(",")","?","*","+","|")
	If strPattern = "" Then strPattern = EmptyString : Exit Function : End If
	' If not an empty string
	For Each Special in vSpecualChars
		strPattern = Replace(strPattern,Special,"\" & Special)
	Next
End Function
'-----------------------------------------------------------------------------------
' Function GetFileLineCountSelect - Returns number of lines int the text file
'-----------------------------------------------------------------------------------
  Function GetFileLineCountSelect(strFileName, ByRef vFileLines,strChar1, strChar2, strChar3, nDebug)
    Dim nIndex
	Dim strLine
	Dim objDataFileName
	Dim objRegEx
	Const EMPTY_STRING = "e-m-p-t-y"
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = False
	
	Call NormalizePatterm(strChar1, EMPTY_STRING)
	Call NormalizePatterm(strChar2, EMPTY_STRING)
	Call NormalizePatterm(strChar3, EMPTY_STRING)
	
    strFileWeekStream = ""	
	If objFSO.FileExists(strFileName) Then 
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.OpenTextFile(strFileName)
		If Err.Number <> 0 Then 
			Call TrDebug("GetFileLineCountSelect: ERROR: CAN'T OPEN FILE:", strFileName, objDebug, MAX_LEN, 0, 1)
			On Error Goto 0
			Redim vFileLines(0)
			GetFileLineCountSelect = 0
			Exit Function
		End If
	Else
	    Call TrDebug("GetFileLineCountSelect: ERROR: CAN'T FIND FILE:", strFileName, objDebug, MAX_LEN, 0, nDebug)
		Redim vFileLines(0)
		GetFileLineCountSelect = 0
		Exit Function
	End If 
    Redim vFileLines(0)
	'Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	Call TrDebug("GetFileLineCountSelect: NOW TRYING TO RIGHT INTO AN ARRAY", "", objDebug, MAX_LEN, 0, nDebug)
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		TestLine = strLine
		objRegEx.Pattern =".+"
		If Not objRegEx.Test(TestLine) Then TestLine = EMPTY_STRING
		objRegEx.Pattern ="(^" & strChar1 & ")|(^" & strChar2 & ")|(^" & strChar3 & ")"
		If 	objRegEx.Test(TestLine) = False Then 
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					Call TrDebug("GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex), "", objDebug, MAX_LEN, 0, nDebug)
					nIndex = nIndex + 1
					bResult = True
		End If		
	Loop
	objDataFileName.Close
	Set objDataFileName = Nothing
    GetFileLineCountSelect = nIndex
End Function
'-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function
'---------------------------------------------------------------------------------------
' 	nMode = 2  Then Insert Above
'   nMode = 3  Then Insert Below
' 	nMode = 1  Then Change
'   nMode = 4  Append Line to File
'	Inserts or change Line in Text File at String Number "LineNumber" (count form 1)
'   Function WriteStrToFile(strDirectoryTmp & "\" & strFileLocalSessionTmp, nTime, LineNumber, CHANGE)
'---------------------------------------------------------------------------------------
Function WriteStrToFile(strFile, strNewLine, LineNumber, nMode, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Dim objFSO, objFile
	Const FOR_WRITING = 1
	WriteStrToFile = False
	If LineNumber > 10000 Then objDebug.WriteLine "WriteStrToFile: ERROR: CAN'T OPERATE FILES WITH MORE THEN 1000 STRINGS" End If  
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objFile = objFSO.OpenTextFile(strFile,2,True)
		objFile.WriteLine "Empty"
		If Err.Number = 0 Then 
			objFile.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteStrToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteStrToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteStrToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLIne is number of lines counted like 1,2,...,n
	If nMode = 2 and LineNumber > nFileLine Then nMode = 12 End If
	If nMode = 3 and LineNumber > nFileLine Then nMode = 12 End If	
	If nMode = 1 and LineNumber > nFileLine Then nMode = 12 End If
	If nMode = 4 Then LineNumber = nFileLine End If	
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": WriteStrToFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	Select Case nMode
			Case 1 																		' - CHANGE REQUESTED LINENUMBER
					vFileLine(LineNumber - 1) = strNewLine
					If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 2
					Redim vvFileLine(nFileLine + 1)
					For i = 0 to LineNumber - 2
						vvFileLine(i) = vFileLine(i)
					Next
					vvFileLine(LineNumber - 1) = strNewLine
					For i = LineNumber to nFileLine
						vvFileLine(i) = vFileLine(i-1)
					Next
					nFileLine = nFileLine + 1
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 3 ' - Insert After
					Redim vvFileLine(nFileLine + 1)
					For i = 0 to LineNumber - 1
						vvFileLine(i) = vFileLine(i)
					Next
					vvFileLine(LineNumber) = strNewLine
					For i = LineNumber + 1 to nFileLine
						vvFileLine(i) = vFileLine(i-1)
					Next
					nFileLine = nFileLine + 1
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 4 ' - Append string to file
			        Set objFile = objFSO.OpenTextFile(strFile,8,True)
					objFile.WriteLine strNewLine
					objFile.Close
					WriteStrToFile = True
			Case 12
					Redim vvFileLine(LineNumber)
					For i = 0 to nFileLine - 1
						vvFileLine(i) = vFileLine(i)
					Next
					For i = nFileLine to LineNumber - 2
						vvFileLine(i) = " "
					Next
					vvFileLine(LineNumber - 1) = strNewLine
					nFileLine = LineNumber
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
	End Select
	Set objFSO = Nothing
	Set objFile = Nothing
End Function
'#######################################################################
 ' Function WriteArrayToFile - Returns number of lines int the text file
 ' nMode = 1  Then Rewire all File content
 ' nMode = 2  Then Append
 ' Creates File if it doesn't exists
 '#######################################################################
 Function WriteArrayToFile(strFile,vFileLine, nFileLine,nMode,nDebug)
    Dim i, nCount
	Dim strLine
	Dim objDataFileName, objFSO

	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.CreateTextFile(strFile)
		If Err.Number = 0 Then 
			objDataFileName.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	
	Select Case nMode
		Case 1 
			Set objDataFileName = objFSO.OpenTextFile(strFile,2,True)
		Case 2 	
			Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	End Select 

	i = 0
	On Error Resume Next
	Err.Clear
	Do While i < nFileLine
		objDataFileName.WriteLine vFileLine(i)
		If Err.Number <> 0 Then 
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T WRITE TO FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			Exit Do 			
		End If
		i = i + 1
	Loop
	On Error Goto 0
	If i = nFileLine Then WriteArrayToFile = True End If
	objDataFileName.close
	Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------------------
'   Function FindAndReplaceStrInFile(strFile, strFind, strNewLine, nDebug)
'   Search for the First Line which contains "strFind" and Replaces whole Line with "strNewLine"
'---------------------------------------------------------------------------------------
Function FindAndReplaceStrInFile(strFile, strFind, strNewLine, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Const FOR_WRITING = 1
	FindAndReplaceStrInFile = False
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLine is number of lines counted like 1,2,...,n
	LineNumber = GetObjectLineNumber( vFileLine, nFileLine, strFind)
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": FindAndReplaceStrInFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	If LineNumber = 0 Then Exit Function
	vFileLine(LineNumber - 1) = strNewLine
	If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then FindAndReplaceStrInFile = True End If
End Function
'---------------------------------------------------------------------------------------
'   Function FindStrInFile(strFile, strFind, strNewLine, nDebug)
'   Search for the First Line which contains "strFind" and Replaces whole Line with "strNewLine"
'---------------------------------------------------------------------------------------
Function FindStrInFile(strFile, strFind, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Const FOR_WRITING = 1
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLine is number of lines counted like 1,2,...,n
	LineNumber = GetObjectLineNumber( vFileLine, nFileLine, strFind)
	If LineNumber > 0 Then 
		FindStrInFile = LineNumber 
	Else 
		FindStrInFile = 0
	End If
End Function
'----------------------------------------------------------
'   Function Set_IE_obj (byRef objIE)
'----------------------------------------------------------
Function Set_IE_obj (byRef objIE)
	Dim nCount
	Set_IE_obj = False
	nCount = 0
	Do 
		On Error Resume Next
		Err.Clear
		Set objIE = CreateObject("InternetExplorer.Application")
		Select Case Err.Number
			Case &H800704A6 
				wscript.sleep 1000
				nCount = nCount + 1
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				If nCount > 4 Then
					On Error goto 0
					Exit Function
				End If
			Case 0 
				Set_IE_obj = True
				On Error goto 0
				Exit Function
			Case Else 
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				On Error goto 0
				Exit Function
		End Select
	On Error goto 0
	Loop
End Function 
 '------------------------------------------------------------------------------
 ' Function GetTestSeries
 '------------------------------------------------------------------------------
 Function GetTestSeries(strFileName, ByRef vSvc, ByRef vFlavors, nDebug)
    Dim nIndex, i, n, nParam
	Dim strLine
	Dim nGroupSelector, nService, nMaxFlavors, nFlavors
	Dim vFileLines, nFileLines, vService, vLines	
	GetTestSeries = 0
	nGroupSelector = 0
	nMaxFlavors = 0
	nService = GetFileLineCountByGroup(strFileName , vService,"Service","","",0)
	Redim vSvc(nService,2)
	nFileLines = GetFileLineCountSelect(strFileName, vFileLines,"#","NULL","NULL",0)
	'-----------------------------------------------------
	'	COUNT Tasks
	'-----------------------------------------r------------
	For n = 0 to nService - 1
		nFlavors = 0
		nFlavors = GetFileLineCountByGroup(strFileName , vLines,vService(n),"","",0)
		vSvc(n,0) = nFlavors
		vSvc(n,1) = vService(n)
		Call TrDebug ("GetTestSeries: TOTAL TASKS IN CATEGORY: " & vService(n), nFlavors, objDebug, MAX_LEN, 1, nDebug)
	Next
	'-----------------------------------------------------
	'	FIND THE MAXIMUM NUMBER OF ALL TASKS
	'-----------------------------------------------------
	For n = 0 to nService - 1
		nMaxFlavors = MaxQ(nMaxFlavors, vSvc(n,0)) 		
	Next
	Call TrDebug ("GetTestSeries: nMaxFlavor = " & nMaxFlavors, "", objDebug, MAX_LEN, 1, nDebug)
	'-----------------------------------------------------
	'	DEFINE vFlavors Array
	'-----------------------------------------------------
	Redim vFlavors(nService, nMaxFlavors,2)		
	'-----------------------------------------------------
	'	LOAD CATEGORIES PROPERIES
	'-----------------------------------------------------
		nGroupSelector = 0
		For n = 0 to nService - 1
			For nIndex = 0 to nFileLines - 1
				strLine = LTrim(vFileLines(nIndex))
				Select Case Left(strLine,1)
					Case "#"
					Case ""
					Case "["
						Select Case strLine
							Case "[" & vService(n) & "]"
								Call TrDebug ("GetTestSeries: LOAD PROPERTIES FOR [" & vService(n) & "]", "", objDebug, MAX_LEN, 3, nDebug)
								nParam = 0
								nGroupSelector = 1
							Case Else
								nGroupSelector = 0
						End Select
					Case Else	
						If nGroupSelector = 1 Then 
							Call TrDebug ("GetTestSeries:" & strLine, "", objDebug, MAX_LEN, 1, nDebug)					
							If nParam < nMaxFlavors Then 
								vFlavors(n,nParam,0) = Split(strLine,":")(0)
								vFlavors(n,nParam,1) = Split(strLine,":")(1)
								Call TrDebug ("GetTestSeries: Flavors(" & n &  "," & nParam & ",0) = "  & vFlavors(n,nParam,0), "", objDebug, MAX_LEN, 1, 1)	
								Call TrDebug ("GetTestSeries: Flavors(" & n &  "," & nParam & ",1) = "  & vFlavors(n,nParam,1), "", objDebug, MAX_LEN, 1, 1)	
							Else 
								Call TrDebug ("GetTestSeries: ERROR: Flavors(nParam) overflow > " & nMaxFlavors, "", objDebug, MAX_LEN, 1, 1)					
							End If
							nParam = nParam + 1 
						End If
				End Select
			Next
		Next
	GetTestSeries = nService
End Function
'------------------------------------------------
'    MAIN DIALOG FORM 
'------------------------------------------------
Function IE_PromptForInput(ByRef g_objIE, ByRef IE_Window_JustCreated, ByRef IE_Reload_Form, vIE_Scale, nDebug)
	Dim g_objShell, objMonitor, objIE_XLS
'	Dim vFilterList, vPolicerList, vCIR, vCBS
	Dim nInd, Arg4, CFG_Downloaded, YES_NO
	Dim nRatioX, nRatioY, nFontSize_10, nFontSize_12, nButtonX, nButtonY, nA, nB
    Dim intX
    Dim intY
	Dim nCount
	Dim strLogin
	Dim IE_Menu_Bar
	Dim IE_Border
	Dim nLine, nService, nFlavor, nTask
	Dim vvMsg(8,3)
	Dim nMaxFlavors
	Dim objFile, objCfgFile
	Const MAX_PARAM = 40
	Const MAX_BW_PROFILES = 30
	Const TCG_MONITOR = "TCG_monitor"
	Const TCG_TEMPLATE = "Juniper_1G_ACX500_E-LINE_E-LAN_E-Tree_E-Access_TCG_v5.2.xlsx"
	Dim objCell
    Set g_objShell = WScript.CreateObject("WScript.Shell")
	Set objIE_XLS = Nothing
	Call TrDebug ("IE_PromptForInput: OPEN MAIN CONFIG LOADER FORM ", "", objDebug, MAX_LEN, 3, nDebug)				
	'-----------------------------------------------------
	'	FIND THE MAXIMUM NUMBER OF ALL TASKS
	'-----------------------------------------------------
	nMaxFlavors = 0
	For n = 0 to Ubound(vSvc,1) - 1
		nMaxFlavors = MaxQ(nMaxFlavors, vSvc(n,0)) 		
	Next
	Call TrDebug ("IE_PromptForInput: nMaxFlavors = " & nMaxFlavors , "", objDebug, MAX_LEN, 3, nDebug)				
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	'----------------------------------------
	' MAIN VARIABLES OF THE GUI FORM
	'----------------------------------------
	If nRatioX > 1 Then nRatioX = 1 : nRatioY = 1 End If
	Select Case nRatioX
		Case 1
				DiagramFigure = strDirectoryWork & "\Data\mef_diagram_normal.png"
		Case 1600/1920
				DiagramFigure = strDirectoryWork & "\Data\TestBed002.png"
		Case else
				DiagramFigure = strDirectoryWork & "\Data\TestBed002.png"
				nRatioX = 1600/1920
				nRatioX = 900/1080
	End Select
	SettingsFigure = strDirectoryWork & "\data\settings-icon.png"
	BgFigure = strDirectoryWork & "\Data\tech_Background_004.jpg"
	AttentionFigure = strDirectoryWork & "\Data\Attention_icon_30x30.png"
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	nBottom = Round(10 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellH = Round(24 * nRatioY,0)
	LoginTitleW = Round(1000 * nRatioX,0)
	FullTitleW = LoginTitleW + Int(LoginTitleW/4)
	nLeft = Round(20 * nRatioX,0)
	nTab = Round(40 * nRatioX,0)
	CellW = LoginTitleW
	LoginTitleH = Round(40 * nRatioY,0)
	nSaveW = nLeft + nButtonX
	nScoreW = 3 * nSaveW
	nColumn = Int(LoginTitleW/3)	
	nNameW = Int((LoginTitleH - nColumn)/3)
	'------------------------------------------
	'	GET NUMBER OF TASKS LINES
	'------------------------------------------	
	nLine = 25
	FullTitleH = 4 * LoginTitleH + cellH * (nLine) + nButtonY + nBottom
	'
	'   Check if Window Hight Fits the Screen size
	If IE_Reload_Form(0) or IE_Window_JustCreated Then 
		WindowH = IE_Menu_Bar + FullTitleH
		WindowW = IE_Border + FullTitleW
		If WindowW < 300 then WindowW = 300 End If
	Else 
		WindowH = g_objIE.height
		WindowW = g_objIE.width  
	End If 
	'
	'  Set Form Fonts
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)	
	'
	'  Create or reload form content
	If IE_Reload_Form(0) or IE_Window_JustCreated Then 
		'
		'  LOAD CSS classes
		strHTMLline = CopyFileToString(strDirectoryBin & "\main.css")  		
		'------------------------------------------------------
		'   Configuration BUTTON 
		'------------------------------------------------------
		nSpace = 15
		nMenuButtonX = Int(LoginTitleW/4) - 2 * nSpace
		nMenuButtonY = nButtonY
		nTable1_W = Int(LoginTitleW/4) - nSpace
		nTable1_H = FullTitleH - 2 * LoginTitleH - 2 * nSpace
		ButtonAlign = "center"
		CellPadding = 2
		ColSpan = 2
		vEvent = Array("","EMPTY","LOAD","DOWNLOAD","APPLY_FWF","CHECK","EDIT","POPULATE_DNLD","POPULATE_ORIG","POPULATE_ONLINE",_
					   "EMPTY","EMPTY","COPY_CLPBRD_L","COPY_CLPBRD_R","APPLY_BW_PROFILE")
		vBText =  Array("","EMPTY",_
						  "Load Config",_
						  "Save Tested Config",_
						  "Apply Filter",_
						  "Check Config",_
						  "Edit Config", _
						  "Export Tested to TCG",_
						  "Export Original to TCG",_
						  "TCG ON-Line",_
						  "",_
						  "Copy Config to Clipboard",_
						  "LEFT",_
						  "RIGHT",_
						  "Apply Bw Profile")
		vOrder = Array(0,2,3,5,6,1,4,14,10,11,12,13,7,8,9)
		strHTMLline = strHTMLline &_
		"<table border=""0"" cellpadding="""& CellPadding &""" cellspacing=""0"" style="" position: absolute; left: 0px; top: " &  LoginTitleH + nSpace & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: Transparent" &_
		"; height: " & nTable1_H & "px; width: " & nTable1_W & "px;"">" & _
		"<tbody>" 
			For i = 1 to 14
				n = vOrder(i)
				Select Case n
					Case 1
						strHTMLline = strHTMLline & "<tr>" &_
						MainMenuButton(n, vEvent(n),LoginTitleH,"",nMenuButtonX,2 * LoginTitleH,HttpBgColor2,HttpTextColor3, ColSpan,ButtonAlign,vBText(n),vEvent(n),nFontSize_12,"") &_
						"</tr>"				
					Case 10
						strHTMLline = strHTMLline & "<tr>" &_
						MainMenuButton(n, vEvent(n),LoginTitleH,"",nMenuButtonX,nTable1_H - 13 * LoginTitleH - 26 * CellPadding,_
						HttpBgColor2,HttpTextColor3, ColSpan,ButtonAlign,vBText(n),vEvent(n),nFontSize_12,"") &_
						"</tr>"				
					Case 11
						strHTMLline = strHTMLline & "<tr>" &_
						MainMenuButton(n, vEvent(n),LoginTitleH,nTable1_W,nMenuButtonX,nMenuButtonY,"Black","Grey", ColSpan,ButtonAlign,vBText(n),vEvent(n),12,"") &_
						"</tr>"
					Case 12
						strHTMLline = strHTMLline & "<tr>" &_
						MainMenuButton(n, vEvent(n),LoginTitleH,nTable1_W/2,nMenuButtonX/2-2,nMenuButtonY,"","", 1,"right",vBText(n),vEvent(n),"","menu_button")
					Case 13
						strHTMLline = strHTMLline &_ 
						MainMenuButton(n, vEvent(n),LoginTitleH,nTable1_W/2,nMenuButtonX/2-2,nMenuButtonY,"","", 1,"left",vBText(n),vEvent(n),"","menu_button") &_
						"</tr>"
					Case Else
						strHTMLline = strHTMLline & "<tr>" &_
						MainMenuButton(n, vEvent(n),LoginTitleH,nTable1_W,nMenuButtonX,nMenuButtonY,"","", ColSpan,ButtonAlign,vBText(n),vEvent(n),"","menu_button") &_
						"</tr>"
				End Select
			Next
		strHTMLline = strHTMLline &	"</tr>" &_					
				"</tbody></table>" &_
				"<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
		'-----------------------------------------------------------------
		' SET THE TITLE OF THE  FORM   		
		'-----------------------------------------------------------------
		nLine = 0
			strHTMLline = strHTMLline &_
			"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
			" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
			"; background-color: " & HttpBgColor5 & "; height: " & LoginTitleH & "px; width: " & FullTitleW & "px;"">" & _
			"<tbody>" & _
			"<tr>" &_
				"<td  style=""border-style: none; background-color: " & HttpBgColor5 & ";""" &_
				"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & FullTitleW - nTab & """>" & _
					"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
					";font-weight: normal;font-style: italic;"">"&_
					"&nbsp;&nbsp;MEF Configuration Loader <span style=""font-weight: bold;"">Ver." & strVersion & "</span></span></p>"&_
				"</td>" &_
					"<td background=""" & SettingsFigure & """ style=""background-repeat: no-repeat; background-position: 50% 50%; background-size: 40px 40px;"&_
					"border-style: none; background-color: " & HttpBgColor5 & ";""" &_
					"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & nTab & """>" & _
					"<button  style='background-color: transparent; border-style: None; width:" &_
					"40;height:40;" &_
					"' name='SETTINGS' onclick=document.all('ButtonHandler').value='SETTINGS_';></button>" & _	
				"</td>" &_			
			"</tr></tbody></table>"	
			'-----------------------------------------------------------------
			' DRAW ROW WITH CONFIGURATION TITLE
			'-----------------------------------------------------------------	
			nTable2_W = FullTitleW - Int(LoginTitleW/4) - nSpace
			nTable2_H = FullTitleH - 2 * LoginTitleH - 2 * nSpace	
			cTitle = "Choose MEF CE2.0 Service Configuration"
			strTitleCell = "<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(nTable2_W/4) & """>" & _
							"<p style=""text-align: center; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							";font-weight: normal;font-style: normal;"">"

			strHTMLline = strHTMLline &_
			"<table border=""0"" cellpadding=""0"" cellspacing=""0""" &_
			"style="" position: absolute; left: " & Int(LoginTitleW/4) & "px; top: " & LoginTitleH + nSpace & "px;" &_
			" border-style: solid; border width: 1px; border-color: " & HttpBgColor2 & ";" &_
			"background-color: " & HttpBgColor2 & "; height: " & LoginTitleH & "px; width: " & nTable2_W & "px;"">" & _
			"<tbody>" & _
			"<tr>" &_
				"<td colspan=""4"" style=""border-style: none; background-color: transparent;""" &_
				"align=""center"" valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW & """>" & _
					"<p><span style="" font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
					";font-weight: normal;font-style: italic;"">" & cTitle & "</span></p>"&_
				"</td>" &_
			"</tr>" &_
			"<tr>" &_
				strTitleCell & "</p></td>"&_
				strTitleCell & "</p></td>"&_
				strTitleCell & "</p></td>"&_
				strTitleCell & "</p></td>"&_					 
			"</tr>"	&_		
			"<tr>" &_
				strTitleCell & "SERVICE</p></td>"&_
				strTitleCell & "TYPE</p></td>"&_
				strTitleCell & "TEST CASE#</p></td>"&_
				strTitleCell & "Use Saved CFG</p></td>"&_					 
			"</tr>"
			'-----------------------------------------------------
			'  SELECT SERVICE NAME
			'-----------------------------------------------------
			InputField = 150
			strHTMLline = strHTMLline &_
			"<tr>"
				strHTMLline = strHTMLline &_
				"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
					"<select name='Input_Param_0' id='Input_Param_0'  class=""regular""" &_
						"style=""width: "& InputField & ";border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
						"position: relative; left:" & nTab & "px; " &_
						"font-family: 'Helvetica'; color: " & HttpTextColor2 & "; " &_
						"background-color:" & HttpBgColor4 & "; font-weight: Normal;"" size='1' " & _
						"onchange=document.all('ButtonHandler').value='Select_0';" &_
						"type=text > " &_
					"</select>" &_
				"</td>"
			'-----------------------------------------------------
			'  SELECT SERVISE FLAVOR
			'-----------------------------------------------------
				strHTMLline = strHTMLline &_
				"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & nNameW & """>" &_
					"<select name='Input_Param_1' id='Input_Param_1' class=""regular""" &_
						"style=""width: "& InputField & ";border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
						"position: relative; left:" & nTab & "px; " &_
						"font-family: 'Helvetica'; color: " & HttpTextColor2 & ";" &_
						"background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1' " & _
						"onchange=document.all('ButtonHandler').value='Select_1';" &_
						"type=text > " &_
					"</select>" &_
				"</td>"
			'-----------------------------------------------------
			'  SELECT TEST NUMBER
			'-----------------------------------------------------
				strHTMLline = strHTMLline &_
				"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & nNameW & """>" &_
					"<select name='Input_Param_2' id='Input_Param_2' class=""regular""" &_
						"style=""width: "& InputField & "; border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
						"position: relative; left:" & nTab & "px; " &_
						"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
						" onchange=document.all('ButtonHandler').value='Select_2';" &_
						"type=text > " &_
					"</select>" &_
				"</td>" &_
				"<td style="" border-style: None; background-color: Transparent;"" class=""oa2"" height=""" & LoginTitleH & """ width=""" & nTab & """ align=""middle"">" & _
					"<input type=checkbox name='ConfigLocation' style=""color: " & HttpTextColor2 & ";""" & _
					" onclick=document.all('ButtonHandler').value='CONFIG_SOURCE';" &_
					"value='Original'>" &_
				"</td>"&_
			"</tr>" &_
			"<tr>" &_
				strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>"&_					 
			"</tr>" &_
		"</tbody></table>"
		nLine = nLine + 4					
		'-----------------------------------------------------
		'  SELECT BW PROFILE FILTER
		'-----------------------------------------------------
		strHTMLline = strHTMLline  &_					
				"<input name='BW_Param_1' type='hidden' value='Fake param with 0 index'>" &_
				"<input name='BW_Param_2' type='hidden' value='Fake param with 0 index'>"
			strTitleCell = "<td colspan=""<colspan>"" style="" border-style: none; "" class=""oa2"" height=""" & cellH & """ width=""" & Int(nTable2_W/5) & """>" & _
							"<p style=""text-align: center; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							";font-weight: normal;font-style: normal;"">"
			strInputCell_1_1 = "<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(nTable2_W/5) & """ align=""center"">" &_
								"<table style=""background-color:transparent;border-color:transparent;border-style: None;""><tbody><tr><td>" &_
									"<input name=BW_Param_1 value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
									" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 & ";" &_
									"background-color: transparent; font-weight: Normal;"" AccessKey=i size=6 maxlength=8 " &_
									"type=text >" &_ 
								"</td><td>" &_
									"<p style=""color: " & HttpTextColor2 & "; "">/</p>" &_
								"</td><td>" &_
									"<input name=BW_Param_1 value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
									" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 & ";" &_
									"background-color: transparent; font-weight: Normal;"" AccessKey=i size=6 maxlength=8 " &_
									"type=text >" &_
								"</td></tr></tbody></table>"							
												
			strInputCell_2_1 = "<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(nTable2_W/5) & """ align=""center"">" &_
								"<table style=""background-color:transparent;border-color:transparent;border-style: None;""><tbody><tr><td>" &_
									"<input name=BW_Param_2 value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
									" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 & ";" &_
									"background-color: transparent; font-weight: Normal;"" AccessKey=i size=6 maxlength=8 " &_
									"type=text >" &_ 
								"</td><td>" &_
									"<p style=""color: " & HttpTextColor2 & "; "">/</p>" &_
								"</td><td>" &_
									"<input name=BW_Param_2 value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
									" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 & ";" &_
									"background-color: transparent; font-weight: Normal;"" AccessKey=i size=6 maxlength=8 " &_
									"type=text >" &_
								"</td></tr></tbody></table>"							

			strHTMLline = strHTMLline &_
				"<table border=""0"" cellpadding=""0"" cellspacing=""0""" &_ 
				"style="" position: absolute; left: " & Int(LoginTitleW/4) & "px; top: " & 2 * LoginTitleH + nLine * cellH + 3 * nSpace & "px;" &_
				"border-collapse: collapse; border-style: solid; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: " & HttpBgColor2 & ";"  &_
				"height: " & LoginTitleH & "px; width: " & nTable2_W & "px;"">" & _
				"<tbody>" & _
					"<tr>" &_
						 Replace(strTitleCell,"<colspan>","5") & "</p></td>"&_
					"</tr>" &_
					"<tr>" &_
						 strTitleCell & "FW FILTER</p></td>"&_
						 strTitleCell & "POLICER</p></td>"&_
						 strTitleCell & "CIR / PIR</p></td>"&_
						 strTitleCell & "CBS / PBS</p></td>"&_
						 strTitleCell & "INTERFACES</p></td>"&_					 
					"</tr>" &_
					"<tr>" &_
						 Replace(strTitleCell,"<colspan>","3") & "</p></td>"&_
						 strTitleCell &_
								"</p><table align=""center"" width=""" & Int(nTable2_W/7) & " style=""background-color:transparent;border-color:transparent;border-style: None;""><tbody><tr><td>" &_
									"<p style=""font-size: 10.0pt; color: " & HttpTextColor2 & "; "">Use Defaults Only</p>" &_
								"</td><td width=""20"">" &_
									"<input type=checkbox name='DefBWprofile' id='DefBWprofile' style=""color: " & HttpTextColor2 & ";""" & _
									"onclick=document.all('ButtonHandler').value='SET_DEF_BW_PROFILE';" &_
									"value='Original'>" &_
								"</td></tr></tbody></table>" &_					 
						 "</td>"&_
						 strTitleCell & "A1 A2 B1 C1 B2 C2</p></td>"&_					 
					"</tr>" &_
					"<tr>" &_
						"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & Int(nTable2_W/5) & """>" &_
							"<select name='bw_profile_1' id='bw_profile_1' class=""regular""" &_
							"style="" width: " & Int(nTable2_W/7) & "; border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							";position: relative; left:" & nTab & "px; " &_
							"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
							" onchange=document.all('ButtonHandler').value='CH_FWF_1' type=text > " &_
							"</select>" &_
						"</td>"	&_
						"<td>" &_
							"<select name='Policer_BW_Param_1' id='Policer_BW_Param_1' class=""regular""" &_
							"style=""width: " & Int(nTable2_W/7) & ";border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							";position: relative; left:" & nTab & "px; " &_
							"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
							" onchange=document.all('ButtonHandler').value='CH_POLICER_1' type=text > " &_
							"</select>" &_
						"</td>" &_
						strInputCell_1_1 & "</td>" &_
						strInputCell_1_1 & "</td>" &_
						"<td align=""middle"">" & Flt_Radio(0,0) & Flt_Radio(0,1) & Flt_Radio(1,0) & Flt_Radio(1,1) & Flt_Radio(1,2) & Flt_Radio(1,3) & "</td>" &_
					"</tr>" &_
					"<tr>" &_
						strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>" & strTitleCell & "</p></td>" &_					 
					"</tr>"	&_	
					"<tr>" &_
						"<td style="" border-style: None;"" align=""left"" class=""oa2"" height=""" & cellH & """ width=""" & Int(nTable2_W/5) & """>" &_
							"<select name='bw_profile_2' id='bw_profile_2' class=""regular""" &_
							"style=""width: " & Int(nTable2_W/7) & ";border: none; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							";position: relative; left:" & nTab & "px; " &_
							"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
							" onchange=document.all('ButtonHandler').value='CH_FWF_2' type=text > " &_
							"</select>" &_
						"</td>"	&_
						"<td>" &_
							"<select name='Policer_BW_Param_2' id='Policer_BW_Param_2' class=""regular""" &_
							"style=""width: " & Int(nTable2_W/7) & ";border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							";position: relative; left:" & nTab & "px; " &_
							"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
							" onchange=document.all('ButtonHandler').value='CH_POLICER_2' type=text > " &_
							"</select>" &_
						"</td>" &_
						strInputCell_2_1 & "</td>" &_
						strInputCell_2_1 & "</td>" &_				
						"<td align=""middle"">" & Flt_Radio(0,0) & Flt_Radio(0,1) & Flt_Radio(1,0) & Flt_Radio(1,1) & Flt_Radio(1,2) & Flt_Radio(1,3) & "</td>" &_
					"</tr>" &_	
					"<tr>" &_
						 Replace(strTitleCell,"<colspan>","5") & "</p></td>"&_
					"</tr>" &_				
				"</tbody></table>"

		nLine = nLine + 7
		'-----------------------------------------------------------------
		' NETWORK DIAGRAM
		'-----------------------------------------------------------------
		Dim DiagramW, DiagramH, ImageW, ImageH, BackgroundDivH
		DiagramW = 700 : DiagramH = 250
		ImageW = DiagramW  : ImageH = DiagramH 
		BackgroundDivH = FullTitleH - 3 * LoginTitleH - 5 * nSpace	- nLine * cellH
		htmlEmptyCell = _
				"<td style="" border-color: transparent; border-style: none;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(DiagramW/6) & """></td>"
		
		strHTMLline = strHTMLline &_	
		 "<div id='divDiagram' name='divDiagram' style='color: " & HttpTextColor3 & " ;background-color: " & HttpBgColor2 & "; " &_ 
		 "width: " & nTable2_W  & "px; height: " & BackgroundDivH & "px;" &_
		 "border-style: none; border-color: " & HttpBgColor2 & "; " &_	 
		 "position: absolute; top: " & 2 * LoginTitleH + nLine * cellH + 4 * nSpace - 2 & "px; left: " & Int(LoginTitleW/4) &"px; '>" &_
			"<table border=""0"" cellpadding=""1"" cellspacing=""1"" height=""" & 5 * cellH & """ width=""" & Int(nTable2_W/5) & """ valign=""middle""" &_ 
			"style=""position: relative; top: 0px; left: " & Int(2*nTable2_W/5) & "px;" &_
			"border-collapse: collapse; border-color: transparent; border-style: none; background-color: transparent "">" & _
				"<tbody>" &_
						"<tr height=""" & cellH & """ ><td colspan=""2"" align=""left"" style=""font-size: 10.0pt; color: " & HttpTextColor2 & ";""><u>Default values:</u></td></tr>" &_
						"<tr height=""" & cellH & """ >" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor2 & ";"">MTU:</td>" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor3 & ";"">" & MTU & " bytes</td>" &_
						"</tr>" &_
						"<tr height=""" & cellH & """ >" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor2 & ";"">CBS(min):</td>" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor3 & ";"">" & CBSo & " bytes</td>" &_
						"</tr>" &_
						"<tr height=""" & cellH & """ >" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor2 & ";"">CBS:</td>" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor3 & ";"">" & CBSdef & " bytes</td>" &_
						"</tr>" &_
						"<tr height=""" & cellH & """ >" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor2 & ";"">PBS:</td>" &_
							"<td align=""left"" style=""font-size: 9.0pt;color: " & HttpTextColor3 & ";"">" & PBSdef & " bytes</td>" &_
						"</tr>" &_					
				"</tbody>" &_
			"</table>" &_
		 "</div>"
		
		strHTMLline = strHTMLline &_	
		 "<div id='divDiagram' name='divDiagram' style='color: " & HttpTextColor3 & " ;background-color: transparent; " &_ 
		 "width: " & DiagramW & "px; height: " & DiagramH & "px;" &_
		 "position: absolute; bottom: " & 3 * LoginTitleH & "px; left: " & Int(3 * LoginTitleW/8) & "px; '>" &_
		 "<img src=" & DiagramFigure & " alt="" border=""0"" height=""" & ImageH & """ width=""" & ImageW & """ valign=""middle"" style="" position: relative; left: " & 20 & "px""></img>" &_
		 "</div>"

	   strHTMLline = strHTMLline &_	
		 "<div id='divDiagram' name='divInterfaces' style='color: " & HttpTextColor3 & " ;background-color: transparent; " &_ 
		 "width: " & DiagramW & "px; height: " & DiagramH & "px;" &_
		 "position: absolute; bottom: " & 3 * LoginTitleH & "px; left: " & Int(3 * LoginTitleW/8) & "px; '>" &_ 
			"<table border=""0"" cellpadding=""1"" cellspacing=""1"" height=""" & 9 * cellH & """ width=""" & Int(DiagramW) & """ valign=""middle""" &_ 
			"style="" position: relative; top: 0px; left: 0px;" &_
			"border-collapse: collapse; border-style: none ; background-color: transparent;"">" & _
				"<tbody>" &_
					"<tr>"&_
						htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & UNI_CELL(1,0,"R") & _
					"</tr>" &_
					"<tr>"&_
						htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & _
					"</tr>" &_
					"<tr>" &_
						UNI_CELL(0,0,"L")  & htmlEmptyCell & htmlEmptyCell & UNI_CELL(1,1,"R") & _
					"</tr>" &_
					"<tr>" &_
						htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & _
					"</tr>" &_
					"<tr>" &_
						htmlEmptyCell & UNI_CELL(0,2,"R") & UNI_CELL(1,4,"L") & htmlEmptyCell & _
					"</tr>" &_
					"<tr>" &_
						UNI_CELL(0,1,"L") & htmlEmptyCell & htmlEmptyCell & UNI_CELL(1,2,"R") & _
					"</tr>" &_
					"<tr>" &_
						htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & _
					"</tr>" &_
					"<tr>" &_
						htmlEmptyCell & htmlEmptyCell & htmlEmptyCell & UNI_CELL(1,3,"R") & _
					"</tr>" &_
					"</tbody>" &_
			"</table>"	&_
			"</div>"
		'------------------------------------------------------
		'   BOTTOM INFO BAR 
		'------------------------------------------------------
		strHTMLline = strHTMLline &_
			"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: 0px;" &_
			" border-collapse: collapse; border-style: none; border width: 1px; border-color: None; background-color: " & HttpBgColor5 &_
			"; height: " & LoginTitleH & "px; width: " & FullTitleW & "px;"">" & _
				"<tbody>" & _
					"<tr>" &_
						"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/8) & """>" & _
							"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
							"; font-weight: normal;font-style: italic;"">Platform: </span></p>" &_
						"</td>" & _
						"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/8) & """>" & _
							"<input name=Current_config value='' style=""text-align: left; font-size: " & nFontSize_12 & ".0pt;" &_ 
							" border-style: none; font-family: 'Helvetica'; font-style: italic; color: " & HttpTextColor3 &_
							"; background-color: Transparent; font-weight: Normal;"" AccessKey=i size=30 maxlength=48 " &_
							"type=text > " &_
						"</td>" & _
						"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
							"; font-weight: normal;font-style: italic;"">" & "Loaded Config:</span></p>" &_
						"</td>" & _
						"<td style=""border-style: None; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<input name=Current_config value='' style=""text-align: left; font-size: " & nFontSize_12 & ".0pt;" &_ 
							" border-style: none; font-family: 'Helvetica'; font-style: italic; color: " & HttpTextColor3 &_
							"; background-color: Transparent; font-weight: Normal;"" AccessKey=i size=30 maxlength=48 " &_
							"type=text disabled > " &_
						"</td>" & _
						"<td style=""border-style: None; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor3 & "; width:" &_
							2 * nButtonX & ";height:" & nButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
							"px;' name='EXIT' onclick=document.all('ButtonHandler').value='Cancel';><u>E</u>xit</button>" & _	
						"</td>" & _
			"</tr></tbody></table>"
		'
		'  Create Message Board
		strHTMLline = strHTMLline & CreateMessageDiv(WindowH,nTab)
		MenuH = WindowH - IE_Menu_Bar
		MenuW = LoginTitleW/4
		strHTMLline = strHTMLline & SettingsMenu(LoginTitleH,MenuW,MenuH,nMenuButtonX,nMenuButtonY, HttpBgColor5, "menu_button")
	End If ' End of the building / rebuilding form
	'-----------------------------------------------------------------
	' HTML Form Parameaters
	'-----------------------------------------------------------------
    If IE_Window_JustCreated Then 	
		g_objIE.Offline = True
		g_objIE.navigate "about:blank"
		g_objIE.Document.Body.innerHTML = strHTMLline	
		Do
			WScript.Sleep 100
		Loop While g_objIE.Busy	
		g_objIE.Document.body.Style.FontFamily = "Helvetica"
		g_objIE.Document.body.Style.FontSize = nFontSize_Def
		g_objIE.Document.body.scroll = "no"
		g_objIE.Document.body.Style.overflow = "hidden"
		g_objIE.Document.body.Style.border = "None "
		g_objIE.Document.body.Style.backgroundcolor = "black"
		g_objIE.Document.body.Style.color = HttpTextColor1
		g_objIE.height = WindowH
		g_objIE.width = WindowW  
		g_objIE.document.Title = MAIN_TITLE
		g_objIE.Top = (intY - g_objIE.height)/2
		g_objIE.Left = (intX - g_objIE.width)/2
		g_objIE.Visible = False		
		g_objIE.MenuBar = False
		g_objIE.StatusBar = False
		g_objIE.AddressBar = False
		g_objIE.Toolbar = False
		g_objIE.Visible = False		
	End If  
	If Not IE_Window_JustCreated and IE_Reload_Form(0) Then 	
		g_objIE.height = WindowH
		g_objIE.width = WindowW  
		g_objIE.Document.Body.innerHTML = strHTMLline
		Call FormReload(g_objIE, IE_Reload_Form, False)
	End If 
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title
	'
	'   Settings pannels Zindex
    g_objIE.Document.getElementById("divSettings").style.zindex = "4"
	'
	'  Menu Button Style: Radius
	For i = 1 to 14
		Select Case i 
			Case 1,10
				g_objIE.document.getElementById("Button"&i).Disabled = True			   
				g_objIE.document.getElementById("Button"&i).style.opacity = 0
		End Select
	Next
    '
	'   Platform under test
	Platform = GetValue(PLATFORM_INDEX)
	DUT_Platform = GetValue(PLATFORM_NAME)
	'
	'	SET DEFAULT PARAMETERS
	CFG_Downloaded = False
	YES_NO = False
	CurrentSvc = "Null"
	'
	' LOAD LAST REMEMBERED PROFILE
	nService = vSessionTmp(0)
	nFlavor = vSessionTmp(1)
	nTaskInd = vSessionTmp(2)
	nTask = Split(vFlavors(nService,nFlavor,1),",")(nTaskInd)
    FltCurrent1 = 0
	FltCurrent2 = 0		
	'
	'  Temporary code for debug purposes
    CurrentSvc = nService
	CurrentFlv = nFlavor 
	CurrentTsk = nTask 
    '	
	'  Load service types
	For nInd = 0 to Ubound(vSvc) - 1
		g_objIE.document.getElementById("Input_Param_0").Length = nInd + 1
		g_objIE.document.getElementById("Input_Param_0").Options(nInd).text = vSvc(nInd,1)
		g_objIE.document.getElementById("Input_Param_0").Options(nInd).value = nInd
	Next
	g_objIE.document.getElementById("Input_Param_0").selectedIndex = nService
	'
	'  Load Service Flavores
	For nInd = 0 to nMaxFlavors - 1
		If nInd < vSvc(nService,0) Then 
			g_objIE.document.getElementById("Input_Param_1").Length = nInd + 1
			g_objIE.document.getElementById("Input_Param_1").Options(nInd).text = vFlavors(nService,nInd,0)	
			g_objIE.document.getElementById("Input_Param_1").Options(nInd).value = nInd
		else
		    Exit For
	    End If
	Next
	g_objIE.document.getElementById("Input_Param_1").selectedIndex = nFlavor
	'
	' Load Task List
	For nInd = 0 to MAX_PARAM - 1
	    If nInd <= Ubound(Split(vFlavors(nService,nFlavor,1),",")) Then 
		    g_objIE.document.getElementById("Input_Param_2").Length = nInd + 1
		    g_objIE.document.getElementById("Input_Param_2").Options(nInd).text = Split(vFlavors(nService,nFlavor,1),",")(nInd)
			g_objIE.document.getElementById("Input_Param_2").Options(nInd).value = nInd
		Else 
		    Exit For
		End If
	Next
	g_objIE.document.getElementById("Input_Param_2").selectedIndex = nTaskInd 
	'
    '  Set Initial BW profile parameters	
	For i=0 to 4 
		g_objIE.Document.All("BW_Param_1")(i).Value = "N/A"
        g_objIE.Document.All("BW_Param_2")(i).Value = "N/A"		
	Next
    Call DisableBWParams(g_objIE, "BW_Param_1", True)	
    Call DisableBWParams(g_objIE, "BW_Param_2", True)	    
	
	g_objIE.Document.All("UNI00").Value = vNodes(0,0)
	g_objIE.Document.All("UNI01").Value = vNodes(0,1)
	g_objIE.Document.All("UNI02").Value = vNodes(0,2)
	g_objIE.Document.All("UNI10").Value = vNodes(1,0)
	g_objIE.Document.All("UNI11").Value = vNodes(1,1)
	g_objIE.Document.All("UNI12").Value = vNodes(1,2)
	g_objIE.Document.All("UNI13").Value = vNodes(1,3)	
	g_objIE.Document.All("UNI14").Value = vNodes(1,4)	
	g_objIE.Document.All("Current_config")(0).Value = GetValue(PLATFORM_NAME)
	g_objIE.Document.All("Current_config")(1).Value = "Unknown"
    ' 
	' Set location of the configuration files folder original vs tested
	if g_objIE.Document.All("ConfigLocation").Checked then	
		SourceFolder = strDirectoryConfig & "\Tested"
		Arg4 = "tested"
	Else
		SourceFolder = strDirectoryConfig
		Arg4 = "original"
    end if
	'
	' Set initial BW PROFILE
	strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
	strConfigFileL = SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL
    Call GetFileLineCountSelect(strConfigFileL,vConfigFileLeft,"","","",0)
	Call LoadBwProfiles_New(g_objIE, "bw_profile_1", "BW_Param_1",FltCurrent1)
	Call LoadBwProfiles_New(g_objIE, "bw_profile_2", "BW_Param_2",FltCurrent2)	
    '
	'  Unhide newly created window
	If IE_Window_JustCreated Then 
		Call IE_Unhide(g_objIE)
		IE_Window_JustCreated = False
	End If

'	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
'	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
'   strPID = ""
'	For Each strLine in vCmdOut
'	   If InStr(strLine,"iexplore.exe") then strPID = Split(strLine,""",""")(1)
'   Next
'	g_objShell.AppActivate strPID	
	'----------------------------------------------------
	'  START MAIN CYCLE OF THE INPUT FORM
	'----------------------------------------------------
	g_objIE.Document.All("ButtonHandler").Value = "CHECK"
    Do
        On Error Resume Next
			If g_objIE.width <> WindowW Then g_objIE.width = WindowW End If
			If g_objIE.height <> WindowH Then g_objIE.height = WindowH End If
			Err.Clear
            szNothing = g_objIE.Document.All("ButtonHandler").Value
            if Err.Number <> 0 then szNothing = "ExitNow"  : IE_PromptForInput = -1 : End If           
        On Error Goto 0    
        Select Case szNothing
		    Case "CONFIG_SOURCE"
						if g_objIE.Document.All("ConfigLocation").Checked then
						    SourceFolder = strDirectoryConfig & "\Tested"
							Arg4 = "tested"
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf") Then
								vvMsg(0,0) = "CONFIGURATION: " &  vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "HASN'T BEEN TESTED YET."                    : vvMsg(1,1) = "normal" 	: vvMsg(1,2) =  HttpTextColor2
								vvMsg(2,0) = "USE ORIGINAL CONFIGURATION INSTEAD"         : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
								Call IE_MSG_Internal(vIE_Scale, "Can't find configuration",vvMsg, 3, g_objIE)
							    SourceFolder = strDirectoryConfig
							    Arg4 = "original"
								g_objIE.Document.All("ConfigLocation").Checked = False
							End If
						Else
							SourceFolder = strDirectoryConfig
							Arg4 = "original"
						end if
						'
						'   Uncheck box for default values
						If g_objIE.Document.All("ConfigLocation").Checked Then g_objIE.Document.All("DefBWprofile").checked = False
						g_objIE.Document.All("ButtonHandler").Value = "CHECK"
			Case "Select_0"
			           g_objIE.Document.All("ButtonHandler").Value = "None"
			            Do
							nNewService = g_objIE.document.getElementById("Input_Param_0").selectedIndex
							If nNewService > Ubound(vSvc)-1 Then nNewService = Ubound(vSvc)-1 : g_objIE.document.getElementById("Input_Param_0").selectedIndex = nNewService End If
							nNewFlavor = 0
							nNewTaskInd = 0
							nNewTask = Split(vFlavors(nNewService,0,1),",")(0)
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nNewService,1) & "\" & vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform & "-l.conf") Then
							    If g_objIE.Document.All("ConfigLocation").Checked Then 
									vvMsg(0,0) = "CONFIGURATION: " &  vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
									vvMsg(1,0) = "HASN'T BEEN TESTED YET."                                                                                          : vvMsg(1,1) = "normal" 	: vvMsg(1,2) =  HttpTextColor2
									vvMsg(2,0) = "ORIGINAL CONFIGURATION WILL BE USED"     : vvMsg(2,1) = "bold" 	                                                : vvMsg(2,2) =  HttpTextColor1
									Call IE_MSG_Internal(vIE_Scale, "Can't find configuration",vvMsg, 3, g_objIE)
									g_objIE.document.getElementById("Input_Param_0").selectedIndex = nService
									Exit Do
								Else 
									vvMsg(0,0) = "CONFIGURATION FILE FOR: " &  vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform : vvMsg(0,1) = "normal" 	: vvMsg(0,2) =  HttpTextColor2
									vvMsg(1,0) = "WAS NOT FOUND."                                                                                   : vvMsg(1,1) = "normal" 	: vvMsg(1,2) =  HttpTextColor2
									vvMsg(2,0) = "MAKE SURE THAT CONFIGURATION FILE EXISTS AND HAS THE RIGHT NAME."                                     : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
									Call IE_MSG_Internal(vIE_Scale, "Can't find file",vvMsg, 3, g_objIE)
									g_objIE.document.getElementById("Input_Param_0").selectedIndex = nService
									Exit Do
								End If 
							End If
							nService = nNewService
							nFlavor = nNewFlavor
							nTaskInd = nNewTaskInd
							nTask = nNewTask
							'  Load Service Flavores
							nInd = 0 
							Do While  nInd < vSvc(nService,0)
								g_objIE.document.getElementById("Input_Param_1").Length = nInd + 1
								g_objIE.document.getElementById("Input_Param_1").Options(nInd).text = vFlavors(nService,nInd,0)	
								nInd = nInd + 1
							Loop
							g_objIE.document.getElementById("Input_Param_1").selectedIndex = 0
							' Load Task List
							nInd = 0
							Do While nInd <= Ubound(Split(vFlavors(nService,nFlavor,1),","))
									g_objIE.document.getElementById("Input_Param_2").Length = nInd + 1
									g_objIE.document.getElementById("Input_Param_2").Options(nInd).text = Split(vFlavors(nService,nFlavor,1),",")(nInd)
									nInd = nInd + 1
							Loop
							g_objIE.document.getElementById("Input_Param_2").selectedIndex = 0
							FltCurrent1 = 0
							FltCurrent2 = 0
							'--------------------------------------
							' LOAD BW PROFILES
							'--------------------------------------
							g_objIE.Document.All("ButtonHandler").Value = "CHECK"
							Exit Do
						Loop       
            Case "Select_1" 
			            g_objIE.Document.All("ButtonHandler").Value = "None"
                        Do			
						    '-------------------------------------------------------
							'   CHECK THAT CONFIGURATION FILE EXISTS
							'-------------------------------------------------------
							nNewService = nService
							nNewFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex
							If nNewFlavor > vSvc(nNewService,0) - 1 Then nNewFlavor = vSvc(nNewService,0) - 1 : g_objIE.document.getElementById("Input_Param_1").selectedIndex = nNewFlavor End If
							nNewTaskInd = 0
							nNewTask = Split(vFlavors(nNewService,nNewFlavor,1),",")(0)
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nNewService,1) & "\" & vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform & "-l.conf") Then
							    If g_objIE.Document.All("ConfigLocation").Checked Then 
									vvMsg(0,0) = "CONFIGURATION: " &  vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
									vvMsg(1,0) = "HASN'T BEEN TESTED YET."                                                                 : vvMsg(1,1) = "normal" 	: vvMsg(1,2) =  HttpTextColor2
									vvMsg(2,0) = "ORIGINAL CONFIGURATION WILL BE USED"     : vvMsg(2,1) = "bold" 	                       : vvMsg(2,2) =  HttpTextColor1
									Call IE_MSG_Internal(vIE_Scale, "Can't find configuration",vvMsg, 3, g_objIE)
									g_objIE.document.getElementById("Input_Param_1").selectedIndex = nFlavor
									Exit Do
								Else 
									vvMsg(0,0) = "CONFIGURATION FILE FOR: " &  vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform : vvMsg(0,1) = "Normal" 	: vvMsg(0,2) =  HttpTextColor2
									vvMsg(1,0) = "WAS NOT FOUND."                                                                                   : vvMsg(1,1) = "Normal" : vvMsg(1,2) =  HttpTextColor2
									vvMsg(2,0) = "MAKE SURE THAT CONFIGURATION FILE HAS RIGHT NAME."                                                : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
									Call IE_MSG_Internal(vIE_Scale, "Can't find file",vvMsg, 3, g_objIE)
									g_objIE.document.getElementById("Input_Param_1").selectedIndex = nFlavor
									Exit Do
								End If 
							End If
							nService = nNewService
							nFlavor = nNewFlavor
							nTaskInd = nNewTaskInd
							nTask = nNewTask
							FltCurrent1 = 0
							FltCurrent2 = 0							
							' Reload Task List
							nInd = 0
							Do While nInd <= Ubound(Split(vFlavors(nService,nFlavor,1),","))
									g_objIE.document.getElementById("Input_Param_2").Length = nInd + 1
									g_objIE.document.getElementById("Input_Param_2").Options(nInd).text = Split(vFlavors(nService,nFlavor,1),",")(nInd)
									nInd = nInd + 1
							Loop
							g_objIE.document.getElementById("Input_Param_2").selectedIndex = 0
							'--------------------------------------
							' LOAD BW PROFILES
							'--------------------------------------
							g_objIE.Document.All("ButtonHandler").Value = "CHECK"
							Exit Do
						Loop
					
            Case "Select_2" 
			            g_objIE.Document.All("ButtonHandler").Value = "None"
                        Do	
						    '-------------------------------------------------------
							'   CHECK THAT CONFIGURATION FILE EXISTS
							'-------------------------------------------------------
							nNewService = nService
							nNewFlavor = nFlavor
							nNewTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
					        nNewTask = g_objIE.document.getElementById("Input_Param_2").options(nNewTaskInd).text
							If LTrim(nNewTask) = "" Then 
							    g_objIE.document.getElementById("Input_Param_2").selectedIndex = 0 
								nNewTaskInd = 0 
								nNewTask = g_objIE.document.getElementById("Input_Param_2").options(nNewTaskInd).text
							End If
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nNewService,1) & "\" & vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform & "-l.conf") Then
							    If g_objIE.Document.All("ConfigLocation").Checked Then 
									vvMsg(0,0) = "CONFIGURATION: " &  vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
									vvMsg(1,0) = "HASN'T BEEN TESTED YET."                    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
									vvMsg(2,0) = "ORIGINAL CONFIGURATION WILL BE USED"     : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
									Call IE_MSG_Internal(vIE_Scale, "Can't find configuration",vvMsg, 3, g_objIE)
									g_objIE.document.getElementById("Input_Param_2").selectedIndex = nTaskInd
									Exit Do
								Else 
									vvMsg(0,0) = "CONFIGURATION FILE FOR: " &  vFlavors(nNewService, nNewFlavor,0) & "-" & nNewTask & "-" & Platform : vvMsg(0,1) = "normal" 	: vvMsg(0,2) =  HttpTextColor2
									vvMsg(1,0) = "WAS NOT FOUND."                                                                                   : vvMsg(1,1) = "normal" : vvMsg(1,2) =  HttpTextColor2
									vvMsg(2,0) = "MAKE SURE THAT CONFIGURATION FILE EXISTS AND HAS THE RIGHT NAME."                                                : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
									Call IE_MSG_Internal(vIE_Scale, "Can't find file",vvMsg, 3, g_objIE)
									g_objIE.document.getElementById("Input_Param_2").selectedIndex = nTaskInd
									Exit Do
								End If 
							End If
							nService = nNewService
							nFlavor = nNewFlavor
							nTaskInd = nNewTaskInd
							nTask = nNewTask
							FltCurrent1 = 0
							FltCurrent2 = 0
							'--------------------------------------
							' LOAD BW PROFILES
							'--------------------------------------
							g_objIE.Document.All("ButtonHandler").Value = "CHECK"
						    Exit Do
						Loop
			Case "CH_FWF_1"
			            FltCurrent1 = g_objIE.document.getElementById("bw_profile_1").selectedIndex
			            Call SelectBwProfiles_New(g_objIE, "bw_profile_1", "BW_Param_1",FltCurrent1)
						g_objIE.Document.All("ButtonHandler").Value = "None"
						Call EnableBwProfiles(g_objIE)
			Case "CH_FWF_2"
						FltCurrent2 = g_objIE.document.getElementById("bw_profile_2").selectedIndex			
			            Call SelectBwProfiles_New(g_objIE, "bw_profile_2", "BW_Param_2",FltCurrent2 )
						g_objIE.Document.All("ButtonHandler").Value = "None"
						Call EnableBwProfiles(g_objIE)	
            Case "CH_POLICER_1"
						Call SelectPolicer_New(g_objIE, "BW_Param_1")
						g_objIE.Document.All("ButtonHandler").Value = "None"
						Call EnableBwProfiles(g_objIE)	
            Case "CH_POLICER_2"			
						Call SelectPolicer_New(g_objIE, "BW_Param_2")
						g_objIE.Document.All("ButtonHandler").Value = "None"
						Call EnableBwProfiles(g_objIE)				
			Case "SET_DEF_BW_PROFILE"
						FltCurrent1 = g_objIE.document.getElementById("bw_profile_1").selectedIndex
						FltCurrent2 = g_objIE.document.getElementById("bw_profile_2").selectedIndex			
			            Call SelectBwProfiles_New(g_objIE, "bw_profile_1", "BW_Param_1",FltCurrent1)
			            Call SelectBwProfiles_New(g_objIE, "bw_profile_2", "BW_Param_2",FltCurrent2)						
						g_objIE.Document.All("ButtonHandler").Value = "None"
						Call EnableBwProfiles(g_objIE)
			Case "APPLY_FWF"
			            g_objIE.Document.All("ButtonHandler").Value = "None"
						Do
				            If Not SecureCRT_Installed Then 
							   Exit Do
    						End If 
							If g_objIE.document.getElementById("bw_profile_1").Disabled and g_objIE.document.getElementById("bw_profile_2").Disabled Then 
								vvMsg(0,0) = "CAN'T APPLY FW FILTER: " 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "ACTION IS NOT AVAILABLE"    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								vvMsg(2,0) = "FOR THIS TEST CONFIGURATION"    : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
								Call IE_MSG_Internal(vIE_Scale, "Applying BW Profile?",vvMsg, 3, g_objIE)
					    		Exit Do
							End If 
							nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex
							nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex
							nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
							nTask = g_objIE.document.getElementById("Input_Param_2").options(nTaskInd).text
							Call TrDebug ("IE_PromptForInput: Current Config: " & CurrentSvc & " " & CurrentFlv & " " & CurrentTsk, "", objDebug, MAX_LEN, 1, 1)						
							Call TrDebug ("IE_PromptForInput: Applyng Config: " & nService & " " & nFlavor & " " & nTask, "", objDebug, MAX_LEN, 1, 1)						
							If CStr(nService) <> CStr(CurrentSvc) or CStr(nFlavor) <> CStr(CurrentFlv) or CStr(nTask) <> CStr(CurrentTsk) Then 
								vvMsg(0,0) = "CAN'T APPLY FW FILTER:" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "LOAD CONFIGURATION FIRST!"    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								Call IE_MSG_Internal (vIE_Scale, "Applying BW Profile?", vvMsg, 2, g_objIE)
								Exit Do
							End If
							'---------------------------------------------------------
							' RUN TELNET SCRIPT
							'---------------------------------------------------------
							strFilter = ApplyFilterList(g_objIE, FltCurrent1, FltCurrent2)
							If strFilter = "" Then 
								vvMsg(0,0) = "There is no Filters to apply:" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "Exit..."							    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								Call IE_MSG_Internal (vIE_Scale, "Applying BW Profile?", vvMsg, 2, g_objIE)
								Exit Do
							End If
							If g_objIE.Document.All("DefBWprofile").checked Then 
								strFilter = strFilter & ApplyBwProfile_All(g_objIE)
								Call WriteStrToFile(strDirectoryWork & "\appl_filter.txt",strCRTexe & strFilter & _
									" /ARG S /ARG " & strFileSettings &_
									" /ARG D /ARG " & strDirectoryWork &_									
									" /SCRIPT " &  VBScript_FLT_and_BWP_Apply, 1, 4, 0)		
								g_objShell.run strCRTexe & strFilter & _
									" /ARG S /ARG " & strFileSettings &_
									" /ARG D /ARG " & strDirectoryWork &_									
									" /SCRIPT " & VBScript_FLT_and_BWP_Apply,nWindowState
									Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)						
									Call TrDebug ("IE_PromptForInput: " & strFilter, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & VBScript_FLT_and_BWP_Apply, "",objDebug, MAX_LEN, 1, 1)						
								Exit Do
								
							End If ' VBScript_FLT_and_BWP_Apply

							Call WriteStrToFile(strDirectoryWork & "\appl_filter.txt",strCRTexe & strFilter & _
								" /ARG S /ARG " & strFileSettings &_
								" /ARG D /ARG " & strDirectoryWork &_									
								" /SCRIPT " &  VBScript_FWF_Apply, 1, 4, 0)		
							
							g_objShell.run strCRTexe & strFilter & _
								" /ARG S /ARG " & strFileSettings &_
								" /ARG D /ARG " & strDirectoryWork &_									
								" /SCRIPT " & VBScript_FWF_Apply,nWindowState
								Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & strFilter, "", objDebug, MAX_LEN, 1, 1)						
							Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & VBScript_FWF_Apply, "",objDebug, MAX_LEN, 1, 1)						
							Exit Do
						Loop
			
			Case "APPLY_BW_PROFILE"
			            g_objIE.Document.All("ButtonHandler").Value = "None"
						Do
				            If Not SecureCRT_Installed Then 
							   Exit Do
    						End If 
							'
							'  Exit if filters are disabled for this configuration
							If g_objIE.document.getElementById("bw_profile_1").Disabled and g_objIE.document.getElementById("bw_profile_2").Disabled Then 
								vvMsg(0,0) = "CAN'T APPLY BW PROFILE: " 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "ACTION IS NOT AVAILABLE"    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								vvMsg(2,0) = "FOR THIS TEST CONFIGURATION"    : vvMsg(2,1) = "bold" 	: vvMsg(2,2) =  HttpTextColor1
								Call IE_MSG_Internal(vIE_Scale, "Applying BW Profile?",vvMsg, 3, g_objIE)
					    		Exit Do
							End If
							nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex
							nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex
							nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
							nTask = g_objIE.document.getElementById("Input_Param_2").options(nTaskInd).text
							Call TrDebug ("IE_PromptForInput: Current Config: " & CurrentSvc & " " & CurrentFlv & " " & CurrentTsk, "", objDebug, MAX_LEN, 1, 1)						
							Call TrDebug ("IE_PromptForInput: Applyng Config: " & nService & " " & nFlavor & " " & nTask, "", objDebug, MAX_LEN, 1, 1)						
							If CStr(nService) <> CStr(CurrentSvc) or CStr(nFlavor) <> CStr(CurrentFlv) or CStr(nTask) <> CStr(CurrentTsk) Then 
								vvMsg(0,0) = "CAN'T APPLY FW FILTER:" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "LOAD CONFIGURATION FIRST!"    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								Call IE_MSG_Internal (vIE_Scale, "Applying BW Profile?", vvMsg, 2, g_objIE)
								Exit Do
							End If
							'
							'   Validate values entered under BW profile: CIR/PIR CBS/EBS
							If Not ValidateBWProfile(g_objIE) Then 
								vvMsg(0,0) = "CAN'T APPLY BW PROFILE:" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "CHECK VALUES USED FOR CIR/PIR CBS/EBS"    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								Call IE_MSG_Internal (vIE_Scale, "Applying BW Profile?", vvMsg, 2, g_objIE)
								Exit Do
							End If 
							'---------------------------------------------------------
							' RUN TELNET SCRIPT
							'---------------------------------------------------------
							strFilter = ApplyBwProfile(g_objIE)
							If strFilter = "" Then 
								vvMsg(0,0) = "There is no BW Profile to apply:" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
								vvMsg(1,0) = "Exit..."							    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								Call IE_MSG_Internal (vIE_Scale, "Applying BW Profile?", vvMsg, 2, g_objIE)
								Exit Do
							End If
							Call WriteStrToFile(strDirectoryWork & "\appl_filter.txt",strCRTexe & strFilter & _
								" /ARG S /ARG " & strFileSettings &_
								" /ARG D /ARG " & strDirectoryWork &_									
								" /SCRIPT " & VBScript_BWP_Apply, 1, 4, 0)		
							
							'Exit Do ' Remove after debug completion 
							g_objShell.run strCRTexe & strFilter & _
								" /ARG S /ARG " & strFileSettings &_
								" /ARG D /ARG " & strDirectoryWork &_									
								" /SCRIPT " & VBScript_BWP_Apply,nWindowState
                                Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & strFilter, "", objDebug, MAX_LEN, 1, 1)						
							Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & VBScript_BWP_Apply, "",objDebug, MAX_LEN, 1, 1)						
							Exit Do
						Loop
			Case "DOWNLOAD"
						g_objIE.Document.All("ButtonHandler").Value = "None"			
			            Do
				            If Not SecureCRT_Installed Then 
							   Exit Do
    						End If 
    						If CurrentSvc = "Null" Then
								nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex
								nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex 
								nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
								nTask = Split(vFlavors(nService,nFlavor,1),",")(nTaskInd)                   
								vvMsg(0,0) = "CURRENTLY LOADED CONFIG IS UNKNOWN:" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
								vvMsg(1,0) = "DOWNLOAD IT AS:" 						    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
								vvMsg(2,0) = "Service: " & vSvc(nService,1)         	: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor2			
								vvMsg(3,0) = "Configuration:  " & vFlavors(nService, nFlavor,0) & "-" & nTask  	: vvMsg(3,1) = "bold" 	: vvMsg(3,2) = HttpTextColor2
								If IE_CONT(vIE_Scale, "DownLoad configurations?", vvMsg, 4, g_objIE, nDebug) Then 
									CurrentSvc = nService
									CurrentFlv = nFlavor
									CurrentTsk = nTask
								End If
							End If
							If CurrentSvc <> "Null" Then
						        If Not YES_NO Then
									vvMsg(0,0) = "DOWNLOAD CONFIGURATION:" 			    	: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
									vvMsg(1,0) = "Service: " & vSvc(CurrentSvc,1)         	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2			
									vvMsg(2,0) = "Configuration:  " & vFlavors(CurrentSvc, CurrentFlv,0) & "-" & CurrentTsk  	: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
									If Not IE_CONT(vIE_Scale, "Downloading Final Configuration?", vvMsg, 3, g_objIE, nDebug) Then Exit Do
								Else 
								    g_objIE.Document.All("ButtonHandler").Value = "LOAD"			
								End If
								g_objShell.run strCRTexe &_ 
									" /ARG " & vSvc(CurrentSvc,1) &_
									" /ARG " & vFlavors(CurrentSvc, CurrentFlv,0) &_
									" /ARG " & CurrentTsk &_
									" /ARG " & strFileSettings &_
									" /ARG " & strDirectoryWork &_																		
									" /SCRIPT " & VBScript_DNLD_Config, nWindowState
								Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)															
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & vSvc(CurrentSvc,1), "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & vFlavors(CurrentSvc, CurrentFlv,0), "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & CurrentTsk, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & VBScript_DNLD_Config, "",objDebug, MAX_LEN, 1, 1)
								wscript.sleep 5000									
							End If	
							CFG_Downloaded = True
						    Exit Do
						Loop
						YES_NO = False
			Case "CHECK"
						Do
							'--------------------------------------
							' LOAD BW PROFILES
							'--------------------------------------
							strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
							strConfigFileL = SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL
							Call GetFileLineCountSelect(strConfigFileL,vConfigFileLeft,"","","",0)
							Call LoadBwProfiles_New(g_objIE, "bw_profile_1", "BW_Param_1",FltCurrent1)
							Call LoadBwProfiles_New(g_objIE, "bw_profile_2", "BW_Param_2",FltCurrent2)
							
							nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex 
							nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex  
							nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
							nTask = Split(vFlavors(nService,nFlavor,1),",")(nTaskInd)                    
							strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
						    strConfigFileR = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-r.conf"
							'---------------------------------------------------------
							' CHECK IF CONFIGURATION FILE LEFT EXIST. 
							'---------------------------------------------------------
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL) Then 
								Call TrDebug ("IE_PromptForInput: FILE DOESN'T EXIST", "...\" & vSvc(nService,1) & "\" & strConfigFileL, objDebug, MAX_LEN, 1, 1)						
								MsgBox "Configuration File for Node-L doesn't exist" & chr(13) & SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL
								exit Do
							End If
							'---------------------------------------------------------
							' CHECK IF CONFIGURATION FILE RIGHT EXIST
							'---------------------------------------------------------
							If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR) Then 
								Call TrDebug ("IE_PromptForInput: FILE DOESN'T EXIST", "...\" & vSvc(nService,1) & "\" & strConfigFileR, objDebug, MAX_LEN, 1, 1)						
								MsgBox "Configuration File for Node-R doesn't exist" & chr(13) & SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR
								exit Do
							End If
							'---------------------------------------------------------
							' CHECK INTERFACES LEFT
							'---------------------------------------------------------
							For i=0 to 2
								nCount = FindStrInFile(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL, vNodes(0,i), 1)
								If nCount > 0 Then 
									g_objIE.Document.All("UNI0"&i).Value = vNodes(0,i)
									g_objIE.Document.All("UNI_BUTTON0"&i).Value = ""									
									Call TrDebug ("IE_PromptForInput: Looking for UNI in Left cfg: " & vNodes(0,i), "OK " & "(" & nCount & ")", objDebug, MAX_LEN, 1, 1)						
								Else
									g_objIE.Document.All("UNI0"&i).Value = ""
									g_objIE.Document.All("UNI_BUTTON0"&i).Value = "N/A"
									Call TrDebug ("IE_PromptForInput: Looking for UNI in Left cfg:" & vNodes(0,i), "NONE " & "(" & nCount & ")", objDebug, MAX_LEN, 1, 1)															
								End If
							Next
							'---------------------------------------------------------
							' CHECK INTERFACES RIGHT
							'---------------------------------------------------------
							For i=0 to 4
								nCount = FindStrInFile(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR, vNodes(1,i), 1)
								If nCount > 0 Then 
									g_objIE.Document.All("UNI1"&i).Value = vNodes(1,i)
									g_objIE.Document.All("UNI_BUTTON1"&i).Value = ""									
									Call TrDebug ("IE_PromptForInput: Looking for UNI in Right cfg: " & vNodes(1,i), "OK " & "(" & nCount & ")", objDebug, MAX_LEN, 1, 1)						
								Else
									g_objIE.Document.All("UNI1"&i).Value = ""
									g_objIE.Document.All("UNI_BUTTON1"&i).Value = "N/A"									
									Call TrDebug ("IE_PromptForInput: Looking for UNI in Right cfg:" & vNodes(1,i), "NONE " & "(" & nCount & ")", objDebug, MAX_LEN, 1, 1)															
								End If
							Next
							'---------------------------------------------------------
							' ENABLE/DISABLE BwProfile Menu
							'---------------------------------------------------------
							Call EnableBwProfiles(g_objIE)
						Exit Do
						Loop
						g_objIE.Document.All("ButtonHandler").Value = "None"			
			Case "LOAD"
		                g_objIE.Document.All("ButtonHandler").Value = "None"
						Do
				            If Not SecureCRT_Installed Then 
							   Exit Do
    						End If 
							If Not CFG_Downloaded and CurrentSvc <> "Null" Then 
								vvMsg(0,0) = "SAVE CURRENT CONFIGURATION FIRST"  : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
								vvMsg(1,0) = "Service: " & vSvc(CurrentSvc,1) : vvMsg(1,1) = "normal" : vvMsg(1,2) =  HttpTextColor2
								vvMsg(2,0) = "Configuration: " & vFlavors(CurrentSvc, CurrentFlv,0) & "-" & CurrentTsk & "-" & Platform   : vvMsg(2,1) = "bold"     : vvMsg(2,2) =  HttpTextColor2
								If IE_CONT_Internal(vIE_Scale, "Continue?", vvMsg,3, g_objIE, nDebug) Then 
									g_objIE.Document.All("ButtonHandler").Value = "DOWNLOAD"
									YES_NO = True
									Exit Do
								End If
							End If
							nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex 
							nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex  
							nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
							nTask = Split(vFlavors(nService,nFlavor,1),",")(nTaskInd)                   
							vvMsg(0,0) = "LOAD CONFIGURATION:" 						: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
							vvMsg(1,0) = "Service: " & vSvc(nService,1)         	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2			
							vvMsg(2,0) = "Configuration:  " & vFlavors(nService, nFlavor,0) & "-" & nTask  	: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
							If IE_CONT_Internal(vIE_Scale, "Load new configurations?", vvMsg, 3, g_objIE, nDebug) Then 
								strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
								strConfigFileR = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-r.conf"
								'---------------------------------------------------------
								' CHECK IF CONFIGURATION FILE LEFT EXIST. COPY TO BACK FOLDER
								'---------------------------------------------------------
								If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL) Then 
									Call TrDebug ("IE_PromptForInput: FILE DOESN'T EXIST", "...\" & vSvc(nService,1) & "\" & strConfigFileL, objDebug, MAX_LEN, 1, 1)						
									MsgBox "Error: Configuration File for Left Node" & chr(13) & vSvc(nService,1) & "\" & strConfigFileL &  chr(13) & "doesn't Exist"
									g_objIE.Document.All("ButtonHandler").Value = "None"
									exit Do
								End If
								'---------------------------------------------------------
								' CHECK IF CONFIGURATION FILE RIGHT EXIST. COPY TO BACK FOLDER
								'---------------------------------------------------------
								If Not objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR) Then 
									Call TrDebug ("IE_PromptForInput: FILE DOESN'T EXIST", "...\" & vSvc(nService,1) & "\" & strConfigFileR, objDebug, MAX_LEN, 1, 1)						
									MsgBox "Error: Configuration File for Left Node" & chr(13) & vSvc(nService,1) & "\" & strConfigFileR &  chr(13) & "doesn't Exist"
									g_objIE.Document.All("ButtonHandler").Value = "None"
									exit Do
								End If
								'---------------------------------------------------------
								' RUN TELNET SCRIPT
								'---------------------------------------------------------
								g_objShell.run strCRTexe &_ 
									" /ARG " & vSvc(nService,1) &_
									" /ARG " & vFlavors(nService, nFlavor,0) &_
									" /ARG " & nTask &_
									" /ARG " & strFileSettings &_	
									" /ARG " & strDirectoryWork &_																	
									" /ARG " & Arg4 &_
									" /ARG " & GetValue(PLATFORM_NAME) &_								
									" /SCRIPT " & VBScript_Upload_Config, nWindowState
                                Call TrDebug ("IE_PromptForInput: " & strCRTexe, "", objDebug, MAX_LEN, 1, 1)														
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & vSvc(nService,1), "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & vFlavors(nService, nFlavor,0), "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & nTask, "", objDebug, MAX_LEN, 1, 1)						
								Call TrDebug ("IE_PromptForInput: " & " /ARG " & Arg4, "", objDebug, MAX_LEN, 1, 1)													
								Call TrDebug ("IE_PromptForInput: " & " /SCRIPT " & VBScript_Upload_Config, "",objDebug, MAX_LEN, 1, 1)						
								g_objIE.Document.All("Current_config")(1).Value = vSvc(nService,1) & ": " & vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform 
								CurrentSvc = nService
								CurrentFlv = nFlavor
								CurrentTsk = nTask
                                CFG_Downloaded = False							
							End If	
							Exit Do
						Loop
			Case "EDIT"
						Do
						nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex
						nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex
						nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
						nTask = Split(vFlavors(nService,nFlavor,1),",")(nTaskInd)
						strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
						strConfigFileR = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-r.conf"
						Tsys0 = DateDiff("n",D0,Date() & " " & Time()) 
						'---------------------------------------------------------
						' CHECK IF CONFIGURATION FILE LEFT EXIST. COPY TO BACK FOLDER
						'---------------------------------------------------------
						If objFSO.FileExists(strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileL) Then 
								objFSO.CopyFile strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileL, strDirectoryConfig & "\Backup\" & vSvc(nService,1) & "\" & Split(strConfigFileL,".")(0) & "-" & Tsys0 & ".conf", True
						Else
							Call TrDebug ("IE_PromptForInput: FILE DOESN'T EXIST", "...\" & vSvc(nService,1) & "\" & strConfigFileL, objDebug, MAX_LEN, 1, 1)						
							MsgBox "File L doesn't Exist"
							exit Do
						End If
						'---------------------------------------------------------
						' CHECK IF CONFIGURATION FILE RIGHT EXIST. COPY TO BACK FOLDER
						'---------------------------------------------------------
						If objFSO.FileExists(strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileR) Then 
								objFSO.CopyFile strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileR,_
												strDirectoryConfig & "\Backup\" & vSvc(nService,1) & "\" & Split(strConfigFileR,".")(0) & "-" & Tsys0 & ".conf", True
						Else
							Call TrDebug ("IE_PromptForInput: FILE DOESN'T EXIST", "...\" & vSvc(nService,1) & "\" & strConfigFileR, objDebug, MAX_LEN, 1, 1)						
							MsgBox "File R doesn't Exist"
							exit Do
						End If
						'---------------------------------------------------------
						' OPEN CONFIGURATION FILES WITH TEXT EDITOR
						'---------------------------------------------------------
						g_objShell.Run strEditor & " " & strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileL	
						g_objShell.Run strEditor & " "  & strDirectoryConfig & "\" & vSvc(nService,1) & "\" & strConfigFileR
						Exit Do
						Loop
						g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "Cancel"
			    		nService = g_objIE.document.getElementById("Input_Param_0").selectedIndex
						nFlavor = g_objIE.document.getElementById("Input_Param_1").selectedIndex
						nTaskInd = g_objIE.document.getElementById("Input_Param_2").selectedIndex
						Call WriteStrToFile(strFileSessionTmp, nService, 1, 1, 0)
						Call WriteStrToFile(strFileSessionTmp, nFlavor, 2, 1, 0)
						Call WriteStrToFile(strFileSessionTmp, nTaskInd, 3, 1, 0)
						IE_PromptForInput = 0
						g_objIE.quit
						Exit Do
            Case "ExitNow"
                ' Exit the function with value of False
				if IE_PromptForInput => 0 Then g_objIE.Document.All("ButtonHandler").Value = "Nothing"
				Set g_objShell = Nothing
                exit Do		
            Case "Save_FTP" 
                IE_PromptForInput = 1
				g_objIE.Document.All("ButtonHandler").Value = "ExitNow"	
            Case "POPULATE_ONLINE"
			    g_objIE.Document.All("ButtonHandler").Value = "None"
				Call Set_IE_obj (obgIE_XLS)
				obgIE_XLS.Visible = True		
			'   obgIE_XLS.Offline = True
				obgIE_XLS.navigate "https://iometrix-my.sharepoint.com/personal/support_iometrix_com/_layouts/15/WopiFrame2.aspx?sourcedoc=%7B6B145436-FF17-40A8-81C8-7AC74F32CF32%7D&file=Juniper_MX_Summit_CE2.0_TCG.xlsx&action=default"
				Do
					WScript.Sleep 200
				Loop While obgIE_XLS.Busy
			Case "COPY_CLPBRD_L"
			    strConfigFileL = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-l.conf"
				If objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL) Then 
				    Call CopyFileToClipboard(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileL)
				Else 
					vvMsg(0,0) = "Can't find configuration file:" 		         : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
					vvMsg(1,0) = strConfigFileL                                  : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
					vvMsg(2,0) = "in folder: " & SourceFolder & "\" & vSvc(nService,1) & "\"    : vvMsg(2,1) = "normal"      : vvMsg(2,2) = HttpTextColor1
					vvMsg(3,0) = ""							
				   Call IE_MSG_Internal(vIE_Scale, "Error",vvMsg,4, g_objIE)
				End If
				g_objIE.Document.All("ButtonHandler").Value = "None"
				
			Case "COPY_CLPBRD_R"
				strConfigFileR = vFlavors(nService, nFlavor,0) & "-" & nTask & "-" & Platform & "-r.conf"
				If objFSO.FileExists(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR) Then 
					Call CopyFileToClipboard(SourceFolder & "\" & vSvc(nService,1) & "\" & strConfigFileR)
				Else 
					vvMsg(0,0) = "Can't find configuration file:" 		         : vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor2
					vvMsg(1,0) = strConfigFileR                                  : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
					vvMsg(2,0) = "in folder: " & SourceFolder & "\" & vSvc(nService,1) & "\"    : vvMsg(2,1) = "normal"      : vvMsg(2,2) = HttpTextColor1
					vvMsg(3,0) = ""							
				   Call IE_MSG_Internal(vIE_Scale, "Error",vvMsg,4, g_objIE)
				End If
				g_objIE.Document.All("ButtonHandler").Value = "None"
			Case "POPULATE_ORIG"
				vvMsg(0,0) = "WOULD YOU LIKE TO POPULATE :" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
				vvMsg(1,0) = "ALL ORIGINAL CONFIGS" 			    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
				vvMsg(2,0) = "TO TCG XLS TEMPLATES? "           	: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor2			
				If IE_CONT(vIE_Scale, "DownLoad configurations?", vvMsg, 3, g_objIE, nDebug) Then 
					g_objIE.Document.All("ButtonHandler").Value = "POPULATE_ALL"
					SourceCfgFolder = strDirectoryConfig
				Else 
					g_objIE.Document.All("ButtonHandler").Value = "None"
				End If
				
			Case "POPULATE_DNLD"
				vvMsg(0,0) = "WOULD YOU LIKE TO POPULATE :" 		: vvMsg(0,1) = "bold" 	: vvMsg(0,2) =  HttpTextColor1
				vvMsg(1,0) = "ALL DOWNLOADED CONFIGS" 			    : vvMsg(1,1) = "bold" 	: vvMsg(1,2) =  HttpTextColor1
				vvMsg(2,0) = "TO TCG XLS TEMPLATES? "           	: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor2			
				If IE_CONT(vIE_Scale, "DownLoad configurations?", vvMsg, 3, g_objIE, nDebug) Then 
					SourceCfgFolder = strDirectoryConfig & "\Tested"
					g_objIE.Document.All("ButtonHandler").Value = "POPULATE_ALL"
				Else 
					g_objIE.Document.All("ButtonHandler").Value = "None"
				End If
			case "POPULATE_ALL"
				g_objIE.Document.All("ButtonHandler").Value = "None"
        Case "CLICK_ON_SHADOW"
				g_objIE.Document.All("ButtonHandler").Value = "None"
			    Select Case nPressSettings Mod 2
					Case 1
						For MenuX = MenuW to 0 step - 2
							g_objIE.document.getElementById("divSettings").style.right = (MenuX - MenuW) & "px"
						Next
						nPressSettings = nPressSettings + 1
						Call IE_Clear_Window_Shadow(g_objIE)
				End Select
		Case "SETTINGS_"
				g_objIE.Document.All("ButtonHandler").Value = "None"
			    Select Case nPressSettings Mod 2
				    Case 0
							Call IE_Set_Window_Shadow(g_objIE, False)							
							For MenuX = 0 to MenuW step 2
							    g_objIE.document.getElementById("divSettings").style.right = (MenuX - MenuW) & "px"
								'WScript.Sleep 1
							 Next
							 
							 nPressSettings = nPressSettings + 1
					Case Else
							 For MenuX = MenuW to 0 step - 2
							    g_objIE.document.getElementById("divSettings").style.right = (MenuX - MenuW) & "px"
								'WScript.Sleep 1
							 Next
						nPressSettings = nPressSettings + 1
						Call IE_Clear_Window_Shadow(g_objIE)
				End Select
		Case "SET_TOPOLOGY"
		        nMenu = 0
				g_objIE.Document.All("ButtonHandler").Value = "OPEN_SETTINGS"
		Case "SET_CONNECTIVITY"
		        nMenu = 1		 
				g_objIE.Document.All("ButtonHandler").Value = "OPEN_SETTINGS"
		Case "SET_FOLDER"
		        nMenu = 2		
				g_objIE.Document.All("ButtonHandler").Value = "OPEN_SETTINGS"
		Case "SET_CRT"
		        nMenu = 3		
				g_objIE.Document.All("ButtonHandler").Value = "OPEN_SETTINGS"
		Case "SET_OTHER"		
				g_objIE.Document.All("ButtonHandler").Value = "None"
    			For MenuX = MenuW to 0 step - 2
					g_objIE.document.getElementById("divSettings").style.right = (MenuX - MenuW) & "px"
				Next
				nPressSettings = nPressSettings + 1
				Call IE_Clear_Window_Shadow(g_objIE)
		Case "OPEN_SETTINGS"
				g_objIE.Document.All("ButtonHandler").Value = "SET_CLOSE"
				nResult = IE_PromptForSettings_Internal(nMenu, g_objIE, vIE_Scale, vPlatforms, nDebug) 
				Select Case nResult 
					Case 0,1

					Case 2
					    IE_PromptForInput = 2
						g_objIE.Document.All("ButtonHandler").Value = "ExitNow"
					Case -1 					
				End Select 
				Call IE_Clear_Window_Shadow(g_objIE)
		Case "SET_CLOSE"
                g_objIE.Document.All("ButtonHandler").Value = "None"		
			    Select Case nPressSettings Mod 2
					Case 1
						For MenuX = MenuW to 0 step - 2
							g_objIE.document.getElementById("divSettings").style.right = (MenuX - MenuW) & "px"
						Next
						nPressSettings = nPressSettings + 1
						Call IE_Clear_Window_Shadow(g_objIE)
				End Select
		
		End Select
		'
		'  Pause 
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "CHECK"
			Case Else
				WScript.Sleep 100
		End Select
    Loop
End Function
'##############################################################################
'      Function PROMPT FOR FOLDER NAME
'##############################################################################

 Function IE_DialogFolder (vIE_Scale, strTitle, strFolder, vLine, ByVal nLine, nDebug)
	Dim objForm, g_objIE, objShell
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
    Set g_objIE = Nothing
    Set objShell = Nothing
	Dim IE_Menu_Bar
	Dim  IE_Border
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_DialogFolder = -1	
	
	Call Set_IE_obj (g_objIE)
	Set objForm = CreateObject("Shell.Application")

	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy

'    With g_objIE.document.parentwindow.screen
'		intX = .availwidth
'        intY = .availheight
'    End With

	nRatioX = intX/1920
    nRatioY = intY/1080
	LineH   = Round (12 * nRatioY)	
	nHeader = Round(10 * nRatioY,0)
	nTab = Round(20 * nRatioX,0)

	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	WindowW = IE_Border + Round(550 * nRatioX,0)
	WindowH = IE_Menu_Bar + 2 * (5 + nLine) * LineH + nButtonY
	
'   If nDebug = 1 Then MsgBox "intX=" & intX & "   intY=" & intY & "   RatioX=" & nRatioX & "  RatioY=" & nRatioY & "   Cell Width=" & cellW & "  Cell Hight=" & cellH End If

	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left =(intX - WindowW)/2
	strHTMLBody = "<br>"
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
			"<b><p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p></b>" 

	Next
	'---------------------------------------------------
	' SET INPUT FILED AND BUTTUN FOR LCL DIRECTORY
	'---------------------------------------------------
	strHTMLBody = strHTMLBody &_	
				"<input name='UserInput' size='80' maxlength='128' style=""position: absolute; Left: " & nTab & "px; top: " &_
				2 * ( nLine + 2) * LineH & "px; font-size: " & nFontSize_12 &_
				".0pt; border-style: None; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
				"; background-color: " & HttpBgColor2 & "; font-weight: bold;"">"				


	'---------------------------------------------------
	' SET OK and CANCEL
	'---------------------------------------------------
    strHTMLBody = strHTMLBody &_
			    "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nTab & "px; bottom: 4px' name='OK' AccessKey='O' onclick=document.all('ButtonHandler').value='OK';><u>O</u>K</button>" & _
                "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; right:" & nTab & "px; bottom: 4px' name='Cancel' AccessKey='C' onclick=document.all('ButtonHandler').value='Cancel';><u>C</u>ancel</button>" & _
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
	strHTMLBody = strHTMLBody &_
                "<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
				";width:" & 2 * nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & Int(WindowW/2) - nButtonX & "px; bottom: 4px'" &_ 
				"name='Local' AccessKey='L' onclick=document.all('ButtonHandler').value='Local';><u>S</u>elect Folder</button>"


			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.Visible = True

	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.AppActivate g_objIE.document.Title
	g_objIE.Document.All("UserInput").Focus
	g_objIE.Document.All("UserInput").Value = strFolder

	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case strNothing
			Case "Cancel"
				' The user clicked Cancel. Exit the loop
				g_objIE.quit
				Set g_objIE = Nothing
				IE_DialogFolder = False			
				Exit Do
			Case "OK"
				IE_DialogFolder = True
				strFolder = g_objIE.Document.All("UserInput").Value
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
			Case "Local"
				If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  "GetPromptFolder: Local Button Pushed" End If  
				strFolder = g_objIE.Document.All("UserInput").Value
				Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
				If Not objFolder Is Nothing Then
					strFolder = objFolder.self.path
					g_objIE.Document.All("UserInput").Value = strFolder
				End If
				g_objIE.Document.All("ButtonHandler").Value = "None"
		End Select
	Wscript.Sleep 200
Loop
    Set g_objIE = Nothing
    Set objShell = Nothing
	Set objFolder = Nothing
End Function
'---------------------------------------------------------------------------
'   Function Space_html(n)
'---------------------------------------------------------------------------
Function Space_html(n)
Dim Str, i
Str = ""
for i = 1 to n
  Str = Str & "&nbsp;" 
next
Space_html = Str
End Function
'---------------------------------------------------------------------------
'   Function Create_Zero_Configs(vSettings)
'---------------------------------------------------------------------------
Function Create_Zero_Configs(ByRef vSettings, Old_Platform_Index)
Dim objFSO, strCfgGlobal, strCfgRE0, Platform,strDirectoryConfig
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCfgGlobal =       GetValue(27)
	strCfgRE0 =          GetValue(28)
	Platform =           GetValue(14)
	strDirectoryConfig = GetValue(7)
	strHostNameL = 	     Split(GetValue(10),",")(2)
	strHostNameR = 	     Split(GetValue(11),",")(2)
    strLeft_ip	=        GetValue(0)
	strRight_ip =        GetValue(1)
	DUT_Platform =       GetValue(13)
	strGlobalFileL = strCfgGlobal & "-" & Platform & "-l.conf"
	strGlobalFileR = strCfgGlobal & "-" & Platform & "-r.conf"
	strRe0FileL = strCfgRE0 & "-" & Platform & "-l.conf"
	strRe0FileR = strCfgRE0 & "-" & Platform & "-r.conf"
	On Error Resume Next
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgGlobal & "-" & Old_Platform_Index & "-l.conf", True
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgGlobal & "-" & Old_Platform_Index & "-r.conf", True
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgRE0 & "-" & Old_Platform_Index & "-l.conf", True
	objFSO.DeleteFile strDirectoryConfig & "\" & strCfgRE0 & "-" & Old_Platform_Index & "-r.conf", True
	On Error Goto 0
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\" & DUT_Platform & "\" &  strCfgGlobal & ".conf" , _
	                strDirectoryConfig & "\" & strCfgGlobal & "-" & Platform & "-l.conf", True
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\"  & DUT_Platform & "\" &  strCfgGlobal & ".conf" , _
	                strDirectoryConfig & "\" & strCfgGlobal & "-" & Platform & "-r.conf", True
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\"  & DUT_Platform & "\" &  strCfgRE0 & ".conf" , _
	                strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-l.conf", True
	objFSO.CopyFile strDirectoryWork & "\config\zero_config\"  & DUT_Platform & "\" &  strCfgRE0 & ".conf" , _
	                strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-r.conf", True
	Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-l.conf", "host-name","     host-name " & strHostNameL & ";", 0)
    Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-r.conf", "host-name","     host-name " & strHostNameR & ";", 0)
	Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-l.conf", "address","                address " & strLeft_ip & ";", 0)
	Call FindAndReplaceStrInFile(strDirectoryConfig & "\" & strCfgRE0 & "-" & Platform & "-r.conf", "address","                address " & strRight_ip & ";", 0)
End Function
'---------------------------------------------------------------------------
'   Function CheckAddrFormat(strAddr,Preffix required[True/False])
'---------------------------------------------------------------------------
Function CheckAddrFormat(strAddr, bPreffix)
	Do
		nCount = UBound(Split(strAddr,"."))
		If nCount <> 3 Then CheckAddrFormat = False : Exit Do : End If
		If bPreffix Then 
			nCount = UBound(Split(strAddr,"/")) 
			If nCount <> 1 Then CheckAddrFormat = False : Exit Do : End If
		End If
		For i = 0 to 3
		    nOctet = Split(Split(strAddr,"/")(0),".")(i)
		    if Not IsNumeric(nOctet) Then CheckAddrFormat = False : Exit Do : End If
			if i = 0 and (Int(nOctet) < 1 or Int(nOctet) > 255) then CheckAddrFormat = False : Exit Do : End If
			if i > 0 and (Int(nOctet) < 0 or Int(nOctet) > 255) then CheckAddrFormat = False : Exit Do : End If
		Next
        If bPreffix Then 
		    nPrefix = Split(strAddr,"/")(1)
			if Int(nPrefix) < 8 or Int(nPrefix) > 30 then CheckAddrFormat = False : Exit Do : End If
		End If
		CheckAddrFormat = True
		Exit Do
	Loop
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUserSYS
'----------------------------------------------------------------------------------
Function GetScreenUserSYS()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserSYS = strScreenUser
End Function
'----------------------------------------------------------------------------------
'    Function Create_CRT_Sessions(ByRef vSettings,ByRef vOld_Settings )
'----------------------------------------------------------------------------------
Function Create_CRT_Sessions(ByRef vSettings,ByRef vOld_Settings )
Dim strSessionFolder, strSessionNameL, strSessionNameR, strOldSessionFolder, strOldSessionNameL, strOldSessionNameR
Dim  StrWinUser, strLeft_ip, strRight_ip, strCRT_SessionFolder,strDirectoryWork, nIndex
Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strSessionFolder =     Split(GetValue(10),",")(0)
	strSessionNameL =      Split(GetValue(10),",")(1)
	strSessionNameR =      Split(GetValue(11),",")(1)
	strOldSessionFolder =  Split(GetValueOld(10),",")(0)
	strOldSessionNameL =   Split(GetValueOld(10),",")(1)
	strOldSessionNameR =   Split(GetValueOld(11),",")(1)
	strLeft_ip =           Split(GetValueOld(0),"/")(0)
	strRight_ip =          Split(GetValueOld(1),"/")(0)
	strCRT_SessionFolder = GetValueOld(15)
	strDirectoryWork =     GetValue(6)
	StrWinUser = GetScreenUserSYS()
	' Check if folder for your session exists
	If strOldSessionFolder <> strSessionFolder and objFSO.FolderExists(strCRT_SessionFolder & "\" & strOldSessionFolder) Then 
		objFSO.DeleteFolder strCRT_SessionFolder & "\" & strOldSessionFolder, True
	End If
	If Not objFSO.FolderExists(strCRT_SessionFolder & "\" & strSessionFolder) Then 
		objFSO.CreateFolder strCRT_SessionFolder & "\" & strSessionFolder
	End If
	' Check if __DataFolder__.ini file exists in ..\sessions
	If Not objFSO.FileExists(strCRT_SessionFolder & "\__FolderData__.ini") Then 
		objFSO.CopyFile strDirectoryWork & "\config\secureCRT\__FolderData__.ini", strCRT_SessionFolder & "\__FolderData__.ini", True
	End If
	' Check if __DataFolder__.ini has in formation about your session folder
	Call  GetFileLineCountSelect(strCRT_SessionFolder & "\__FolderData__.ini", vFolderData,"", "", "", 0)
	nIndex = GetObjectLineNumber(vFolderData, UBound(vFolderData),"Folder List") - 1
	If strOldSessionFolder <> strSessionFolder and InStr(vFolderData(nIndex),strOldSessionFolder & ":") then 
		vFolderData(nIndex) = Replace(vFolderData(nIndex), strOldSessionFolder & ":","")
	End If
	If Not Instr(vFolderData(nIndex),strSessionFolder  & ":") then 
		vFolderData(nIndex) = vFolderData(nIndex) & strSessionFolder & ":"
		Call WriteArrayToFile(strCRT_SessionFolder & "\__FolderData__.ini",vFolderData, UBound(vFolderData),1,0)
	End If
	' - Delete old session files
	If strOldSessionFolder = strSessionFolder and strOldSessionNameL <> stressionNameL and objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strOldSessionNameL & ".ini") Then
	   objFSO.DeleteFile strCRT_SessionFolder & "\" & strSessionFolder & "\" & strOldSessionNameL & ".ini", True
	End If 
	If strOldSessionFolder = strSessionFolder and strOldSessionNameR <> stressionNameR and objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strOldSessionNameR & ".ini") Then
	   objFSO.DeleteFile strCRT_SessionFolder & "\" & strSessionFolder & "\" & strOldSessionNameR & ".ini", True
	End If 
	' - Create New Session File for Left Node Session
	objFSO.CopyFile strDirectoryWork & "\config\secureCRT\node.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameL & ".ini", True
	Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameL & ".ini", vSessionFile,"", "", "", 0)
	nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{hostname}}") - 1
	vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{hostname}}",strLeft_ip)
	nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{winusername}}") - 1
	vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{winusername}}",StrWinUser)
	nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{Workdirectory}}") - 1
	vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{Workdirectory}}",strDirectoryWork)
	Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameL & ".ini",vSessionFile, UBound(vSessionFile),1,0)
	' - Create New Session File for Right Node Session
	objFSO.CopyFile strDirectoryWork & "\config\secureCRT\node.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameR & ".ini", True
	Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameR & ".ini", vSessionFile,"", "", "", 0)
	nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{hostname}}") - 1
	vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{hostname}}",strRight_ip)
	nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{winusername}}") - 1
	vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{winusername}}",StrWinUser)
	nIndex = GetObjectLineNumber(vSessionFile, UBound(vSessionFile),"{{Workdirectory}}") - 1
	vSessionFile(nIndex) = Replace(vSessionFile(nIndex),"{{Workdirectory}}",strDirectoryWork)
	Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\" & strSessionNameR & ".ini",vSessionFile, UBound(vSessionFile),1,0)
	' Check if __DataFolder__.ini file exists in ..\sessions\SessionName
	If Not objFSO.FileExists(strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini") Then 
		objFSO.CopyFile strDirectoryWork & "\config\secureCRT\__FolderData__.ini", strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini", True
	End If
	' Check if __DataFolder__.ini has information about your new sessions
	Call  GetFileLineCountSelect(strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini", vFolderData,"", "", "", 0)
	nIndex = GetObjectLineNumber(vFolderData, UBound(vFolderData),"Session List") - 1
	If Not Instr(vFolderData(nIndex),strSessionNameL & ":") then 
		vFolderData(nIndex) = vFolderData(nIndex) & strSessionNameL & ":"
	End If
	If Not Instr(vFolderData(nIndex),strSessionNameR & ":") then 
		vFolderData(nIndex) = vFolderData(nIndex) & strSessionNameR & ":"
	End If
	Call WriteArrayToFile(strCRT_SessionFolder & "\" & strSessionFolder & "\__FolderData__.ini",vFolderData, UBound(vFolderData),1,0)	
    Set objFSO = Nothing
End Function
'###################################################################################
' Displays a Message Box with Cancel / Continue buttons                 
'###################################################################################
Function Continue(strMsg, strTitle)
    ' Set the buttons as Yes and No, with the default button
    ' to the second button ("No", in this example)
    nButtons = vbYesNo + vbDefaultButton2
    
    ' Set the icon of the dialog to be a question mark
    nIcon = vbQuestion
    
    ' Display the dialog and set the return value of our
    ' function accordingly
    If MsgBox(strMsg, nButtons + nIcon, strTitle) <> vbYes Then
        Continue = False
    Else
        Continue = True
    End If
End Function 
'--------------------------------------------------------------------
' Function Runs MS CMD Command on local or remote PC
'--------------------------------------------------------------------
Function RunCmd(strHost, strPsExeFolder, ByRef vCmdOut, strCMD, nDebug)	
	Dim nResult, f_objFSO, objShell
	Dim nCmd, stdOutFile, objCmdFile, cmdFile, strRnd,strWork,strPsExec
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	strRnd = My_Random(1,999999)
	stdOutFile = "svc-" & strRnd & ".dat"
	cmdFile = "run-" & strRnd & ".bat"
	If objFSO.FolderExists(strDirectoryTmp) _ 
		Then strWork = strDirectoryTmp _
		Else strWork = objShell.ExpandEnvironmentStrings("%USERPROFILE%")	
	If strHost = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") or strHost = "127.0.0.1" Then 
		strPsExec = ""
	Else 
		strPsExec = strPsExeFolder & "\psexec \\" & strHost & " -s "
	End If
	'-------------------------------------------------------------------
	'       CREATE A NEW TERMINAL SESSION IF REQUIRED
	'-------------------------------------------------------------------
	Set objCmdFile = objFSO.OpenTextFile(strWork & "\" & cmdFile,ForWriting,True)
	Call TrDebug ("COMMAND: ", strPsExec & strCMD & " >" & strWork & "\" & stdOutFile, objDebug, MAX_WIDTH, 1, nDebug)
	objCmdFile.WriteLine strPsExec & strCMD & " >" & strWork & "\" & stdOutFile
	objCmdFile.WriteLine "Exit"
	objCmdFile.close
	objShell.run strWork & "\" & cmdFile,0,True
	Call TrDebug ("BATCH FILE EXECUTED: ", strWork & "\" & cmdFile, objDebug, MAX_WIDTH, 1, nDebug)
	wscript.sleep 100
	'-----------------------------------------
	' READ OUTPUT FILE AND DELETE WHEN DONE
	'-----------------------------------------
	RunCmd = GetFileLineCountSelect(strWork & "\" & stdOutFile, vCmdOut,"NULL","NULL","NULL",nDebug)
	If f_objFSO.FileExists(strWork & "\" & stdOutFile) Then
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & stdOutFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",stdOutFile, objDebug, MAX_WIDTH, 1, 1)
			On Error goto 0
		End If	
	End If
	If f_objFSO.FileExists(strWork & "\" & cmdFile) Then 
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & cmdFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",cmdFile, objDebug, MAX_WIDTH, 1, 1)
			On Error goto 0
		End If		
	End If
	Set f_objFSO = Nothing
	Set objShell = Nothing
	If RunCmd = 0 Then 
		Call TrDebug ("RunCmd: " & strCMD & " ERROR: ", "CAN'T WRITE TO OUTPUT FILE OR EMPTY FILE" , objDebug, MAX_WIDTH, 1, 1)
		Exit Function 
	End If
End Function
'--------------------------------------------------------------
' Function returns a random intiger between min and max
'--------------------------------------------------------------
Function My_Random(min, max)
	Randomize
	My_Random = (Int((max-min+1)*Rnd+min))
End Function
'----------------------------------------------------------------
'   Function MinimizeParentWindow()
'----------------------------------------------------------------
Function MinimizeParentWindow()
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
    wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "n"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function MinimizeParentWindow()
'----------------------------------------------------------------
Function RestoreParentWindow()
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
    wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function FocusToParentWindow(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function FocusToParentWindow(strPID)
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "%"
	wscript.sleep IE_PAUSE
	objShell.AppActivate strPID			
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function GetAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetAppPID(ByRef strPID, strAppName)
Dim objWMI, colItems
Const IE_PAUSE = 70
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Exit Do
		End If 
'		wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'Launcher Ver.'"  WHERE Name = 'iexplore.exe' OR Name = 'wscript.exe'
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("GetMyPID: RESTORE IE WINDOW:", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
			If pUser = strUser then 
				strPID = process.ProcessId
				Call TrDebug ("GetMyPID: ", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
				Call TrDebug ("GetMyPID: ", "Caption: " & process.Caption & ", CSName " & process.CSName & ", Description: " & process.Description & ", Handle: " &  Process.Handle,objDebug, MAX_LEN, 1, 1) 
			GetMyPID = True
				Exit For
			End If
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'-----------------------------------------------------------------------
'   Function GetAppPID(strPID) Returns focus to the parent Window/Form
'-----------------------------------------------------------------------
Function GetFilterList(vConfigFileLeft, ByRef vFilterList, ByRef vPolicerList, ByRef vCIR, ByRef vCBS, nDebug)
Dim nFilter, nPolicer,FW_Start, FoundFilter, FoundPolicer, strLine, nBraces, nBraceFilter
Redim vFilterList(0)
Redim vPolicerList(0)
Redim vCIR(0)
Redim vCBS(0)
    GetFilterList = False
	nFilter = 0
	nBraces = 10000
	For Each strLine in vConfigFileLeft
	    If InStr(strLine,"firewall ")<>0 Then FW_Start = True : nBraces = 0
		If InStr(strLine," {")<>0 Then nBraces = nBraces + 1
		If InStr(strLine," }")<>0 Then nBraces = nBraces - 1
'		If nBraces = 2 Then FoundFilter = False
		If nBraces = 1 Then FoundPolicer = False : FoundFilter = False : End If
		if nBraces = 0 Then Exit For
		If Instr(strLine,"filter ")<>0 and InStr(strLine, " {")<>0 and Instr(strLine,"filter {")=0 Then     		
		    Redim Preserve vFilterList(nFilter + 1)
			Redim Preserve vPolicerList(nFilter + 1)
		    Redim Preserve vCIR(nFilter + 1)
			Redim Preserve vCBS(nFilter + 1)
		    vFilterList(nFilter) = Split(Split(strLine,"filter ")(1)," {")(0)
			Call TrDebug ("GetFilterList: FOUND FILTER:", vFilterList(nFilter),objDebug, MAX_LEN, 1, nDebug)
			GetFilterList = True
			nBraceFilter = 0
			FoundFilter = True
		End If
		If FoundFilter Then
		    '  
			'  Count braces under filter hierarchy level 
			If InStr(strLine," {")<>0 Then nBraceFilter = nBraceFilter + 1
			If InStr(strLine," }")<>0 Then nBraceFilter = nBraceFilter - 1
			If nBraceFilter = 0 Then 
			    If vPolicerList(nFilter) = "" Then 
				    vPolicerList(nFilter) = "N/A"
					vCIR(nFilter) = "N/A"					
					vCBS(nFilter) = ""					
				    Call TrDebug ("GetFilterList: NO POLICER FOUND: ", vPolicerList(nFilter),objDebug, MAX_LEN, 1, nDebug) 
                End If
				FoundFilter = False
				nFilter = nFilter + 1
            End If 				
			'
			'   Validate two-rate policer configured under filter
			If Instr(strLine,"-rate")<>0 Then 
				vPolicerList(nFilter) = Split(Split(strLine,"-rate ")(1),";")(0)
				Call TrDebug ("GetFilterList: FOUND POLICER (key ""-rate""):", vPolicerList(nFilter),objDebug, MAX_LEN, 1, nDebug) 						
			End If
			'
			'   Validate single-rate policer configured under filter
			If Instr(strLine,"policer")<>0 and InStr(strLine,"-policer")=0 and InStr(strLine," {")=0 Then 
				vPolicerList(nFilter) = Split(Split(strLine,"policer ")(1),";")(0)
				Call TrDebug ("GetFilterList: FOUND POLICER (key ""-policer""):", vPolicerList(nFilter),objDebug, MAX_LEN, 1, nDebug)
			End If
		End If
		If Not FoundFilter Then' - Parser is out of Firewall filter hierarchy
			If InStr(strLine,"policer ")<>0 and InStr(strLine," {")<>0 Then 
			    strFilterCount = ""
				FoundPolicer = False
				For nPolicer=0 to UBound(vPolicerList)-1
				   If InStr(strLine,vPolicerList(nPolicer) & " {")<>0 Then 
				        FoundPolicer = True 
						Call TrDebug ("GetFilterList: FOUND SETTINGS FOR POLICER:", vPolicerList(nPolicer),objDebug, MAX_LEN, 1, nDebug)
						If strFilterCount = "" Then strFilterCount = CStr(nPolicer) Else   strFilterCount = strFilterCount & "," & CStr(nPolicer)
					End If
				Next
			End If
			If FoundPolicer Then 
			    vPolicer = Split(strFilterCount,",")
				For each nLine in vPolicer
				    nPolicer = CInt(nLine)
					If InStr(strLine,"committed-information-rate")<>0 Then vCIR(nPolicer) = Trim(Split(Split(strLine,"committed-information-rate ")(1),";")(0))
					If InStr(strLine,"bandwidth-limit")<>0 Then vCIR(nPolicer) = Trim(Split(Split(strLine,"bandwidth-limit ")(1),";")(0))
					If InStr(strLine,"peak-information-rate")<>0 Then vCIR(nPolicer) = vCIR(nPolicer) & "/" & Trim(Split(Split(strLine,"peak-information-rate ")(1),";")(0))
					If InStr(strLine,"committed-burst-size")<>0 Then vCBS(nPolicer) = Trim(Split(Split(strLine,"committed-burst-size ")(1),";")(0))
					If InStr(strLine,"burst-size-limit")<>0 Then vCBS(nPolicer) = Trim(Split(Split(strLine,"burst-size-limit ")(1),";")(0))
					If InStr(strLine,"peak-burst-size")<>0 Then vCBS(nPolicer) = vCBS(nPolicer) & "/" & Trim(Split(Split(strLine,"peak-burst-size ")(1),";")(0))
				Next
			End If
        End If			
	Next
End Function
'-----------------------------------------------------------------------------------------
'      Function Displays a Message with Continue and No Button. Returns True if Continue
'-----------------------------------------------------------------------------------------
 Function IE_CONT (vIE_Scale, strTitle, vLine, ByVal nLine, objIEParent, nDebug)
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim g_objIE, objShell
    Set g_objIE = Nothing
    Set objShell = Nothing
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	IE_CONT = False
	Set objShell = WScript.CreateObject("WScript.Shell")
	Call IE_Hide(objIEParent)
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nRatioX = intX/1920
    nRatioY = intY/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	nTab = Round(20 * nRatioX,0)
	BottomH = Round(10 * nRatioY,0)
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	
 '  If nDebug = 1 Then MsgBox "intX=" & intX & "   intY=" & intY & "   RatioX=" & nRatioX & "  RatioY=" & nRatioY & "   Cell Width=" & cellW & "  Cell Hight=" & cellH End If
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	strHTMLBody = "<br>"
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
			
	Next		
	
    strHTMLBody = strHTMLBody &_
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; left: " & nTab & "px; bottom: " & BottomH & "px' name='Continue' AccessKey='Y' onclick=document.all('ButtonHandler').value='YES';><u>Y</u>ES</button>" & _
								"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; right: " & nTab & "px; bottom: " & BottomH & "px' name='NO' AccessKey='N' onclick=document.all('ButtonHandler').value='NO';><u>N</u>O</button>" & _
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.Visible = True
	IE_Full_AppName = g_objIE.document.Title & " - " & IE_Window_Title
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	'----------------------------------------------------
	'  GET MAIN FORM PID
	'----------------------------------------------------
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & IE_Full_AppName & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, nDebug)
    strMyPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then strMyPID = Split(strLine,""",""")(1)
	     ' Call TrDebug("READ TASK PID:" , strLine, objDebug, MAX_LEN, 1, 1)
    Next
    If strMyPID = "" Then Call GetAppPID(strMyPID, "iexplore.exe")
	objShell.AppActivate strMyPID										
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "NO"
				IE_CONT = False
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
			Case "YES"
				IE_CONT = True
				g_objIE.quit
				Set g_objIE = Nothing
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
		Call IE_UnHide(objIEParent)
End Function
'----------------------------------------------------------------------------------------
'   Function IE_Hide(objIE) changes the visibility of the Window referenced by the objIE
'----------------------------------------------------------------------------------------
Function IE_Hide(byRef objIE)
   if objIE = Null then exit function
   If objIE.Visible then 
        objIE.Visible = False
    End If 
End Function
'----------------------------------------------------------------------------------------
'   Function IE_UnHide(objIE) changes the visibility of the Window referenced by the objIE
'----------------------------------------------------------------------------------------
Function IE_Unhide(byRef objIE)
   if objIE = Null then exit function
   If Not objIE.Visible then 
        objIE.Visible = True
    End If 
End Function
'----------------------------------------------------------------
'   Function GetWinAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetWinAppPID(ByRef strPID, ByRef strParentPID, strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetWinAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("GetWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("GetWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: CMD: " & process.CommandLine, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: ParentPID:" &  Process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug) 			
			Select Case Lcase(strCommandLine)
			    Case "null", "none", ""
					If pUser = strUser then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			    Case Else
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			End Select
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'----------------------------------
' Function UNI_CELL(N,I) 
' N = Node, 0 -> Left, 1 -> Right
' I = UNI # or NNI . NNI Always the last in the list 
'----------------------------------
Function UNI_CELL(N,I,Position)
Const HttpTextColor3 = "#EBEAF7"
Dim nFontSize_10, nLeft, nTop, X, DiagramW, strLine
	nFontSize_10 = 10
	nLeft = 20
	nTop = 2
	DiagramW = 700
	X = N & I
	Select Case Position
		Case "L" 
			Position = "left"
		Case "R" 
			Position = "right"
		Case Else 
			Position = "middle"
	End Select
	strLine = _
		"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(DiagramW/6) & """ align=""" & Position & """>" &_
			"<input name=UNI"& X &" value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
			" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
            ";position: relative; left:" & nLeft & "px; top: " & nTop & "px; " &_			
		    "; background-color: transparent; font-weight: Bold;"" AccessKey=i size=6 maxlength=10 " &_
			"type=text > " &_
			"<input name=UNI_BUTTON"& X &" value='' style=""text-align: center; font-size: " & nFontSize_10 & ".0pt;" &_ 
			" border-style: none; font-family: 'Helvetica'; color: Red" &_
            ";position: relative; left:-" & nLeft & "px; top: " & nTop & "px; " &_						
		    "; background-color: transparent; font-weight: Bold;"" size=1 maxlength=1 " &_
			"type=text > " &_
		"</td>"
	UNI_CELL = strLine
'    Call WriteStrToFile(strDirectoryWork & "\html_diagram.html",strLine, 1, 4, 0)		
End Function
'-----------------------------------------
'   Function Flt_Radio(N,I)
'-----------------------------------------
Function Flt_Radio(N,I)
	Dim X
	X = N & I
	Flt_Radio =	"<input type=radio " &_ 
	                     "id='Flt_Intf"& X &"'" &_ 
						 "name='Flt_Intf"& X &"'" &_ 
						 "style=""color: " & HttpTextColor2 & ";""" & _
						 "value='Original'>"
End Function
'------------------------------------------------
'   Function LoadBwProfiles(g_objIE, html_BW_Profile, html_BW_Params,nFilter)
'------------------------------------------------
Function LoadBwProfiles(ByRef g_objIE, html_BW_Profile, html_BW_Params,nFilter)
Dim CBS, PBS
	If Not GetFilterList(vConfigFileLeft, vFilterList, vPolicerList, vCIR, vCBS, 1) Then 
		g_objIE.document.getElementById(html_BW_Profile).length = 1
		g_objIE.document.getElementById(html_BW_Profile).Options(0).Text = "N/A"
		g_objIE.document.getElementById(html_BW_Profile).SelectedIndex = 0
		g_objIE.document.getElementById(html_BW_Profile).Disabled = True
	Else
		g_objIE.document.getElementById(html_BW_Profile).Disabled = False
		For n = 0 to UBound(vFilterList) - 1
			g_objIE.document.getElementById(html_BW_Profile).length = n + 1
			g_objIE.document.getElementById(html_BW_Profile).Options(n).text = vFilterList(n) 
		Next
		If Int(nFilter) <= n Then g_objIE.document.getElementById(html_BW_Profile).selectedIndex = nFilter
	End If
	If vPolicerList(nFilter) <> "" Then g_objIE.Document.All(html_BW_Params)(0).Value = vPolicerList(nFilter) Else g_objIE.Document.All(html_BW_Params)(0).Value = "N/A" End If
	If InStr(vCIR(nFilter),"/") and vCIR(nFilter) <> "N/A" Then 
		g_objIE.Document.All(html_BW_Params)(1).Value = Split(vCIR(nFilter),"/")(0)
		g_objIE.Document.All(html_BW_Params)(2).Value = Split(vCIR(nFilter),"/")(1)		
	Else 
		g_objIE.Document.All(html_BW_Params)(1).Value = "N/A" 
		g_objIE.Document.All(html_BW_Params)(2).Value = "N/A" 		
	End If
	'
	'  Load CBS/PBS value to the form
	'  Load Either Default value or value from original configuration file
	If g_objIE.Document.All("DefBWprofile").Checked Then 
		If InStr(vCBS(nFilter),"/") and vCBS(nFilter) <> "N/A" Then 
			Select Case ValidateCBS(nFilter)
				Case 1  ' CIR=N and EIR=0
						CBS = CBSdef : PBS = CBS
				Case 2  ' CIR=0 and EIR=N
						CBS = CBSo : PBS = int(CBSo) + int(CBSdef)
				Case 3  ' CIR=N and EIR=N
						CBS = CBSdef : PBS = 2 * CBS				
			End Select
			g_objIE.Document.All(html_BW_Params)(3).Value = CBS 
			g_objIE.Document.All(html_BW_Params)(4).Value = PBS
		Else
			g_objIE.Document.All(html_BW_Params)(3).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(4).Value = "N/A" 	
		End If	
	Else 
		If InStr(vCBS(nFilter),"/") and vCBS(nFilter) <> "N/A" Then 
			g_objIE.Document.All(html_BW_Params)(3).Value = Split(vCBS(nFilter),"/")(0) 
			g_objIE.Document.All(html_BW_Params)(4).Value = Split(vCBS(nFilter),"/")(1) 		
		Else
			g_objIE.Document.All(html_BW_Params)(3).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(4).Value = "N/A" 	
		End If	
	End if 
	For i = 0 to 1
		g_objIE.Document.All("Flt_Intf0" & i )(0).Disabled = False
		g_objIE.Document.All("Flt_Intf0" & i )(1).Disabled = False
		g_objIE.Document.All("Flt_Intf0" & i )(0).Checked = False
		g_objIE.Document.All("Flt_Intf0" & i )(1).Checked = False
	Next 
	For i = 0 to 3
	    g_objIE.Document.All("Flt_Intf1" & i )(0).Disabled = False
		g_objIE.Document.All("Flt_Intf1" & i )(1).Disabled = False
	    g_objIE.Document.All("Flt_Intf1" & i )(0).Checked = False
		g_objIE.Document.All("Flt_Intf1" & i )(1).Checked = False
	Next 	
End Function
'------------------------------------------------
'   Function SelectBwProfiles(g_objIE, html_BW_Profile, html_BW_Params)
'------------------------------------------------
Function SelectBwProfiles(ByRef g_objIE, html_BW_Profile, html_BW_Params)
Dim nFilter
	nFilter = g_objIE.document.getElementById(html_BW_Profile).selectedIndex
	If LTrim(g_objIE.document.getElementById(html_BW_Profile).Options(nFilter).Text) = "" Then 
		g_objIE.document.getElementById(html_BW_Profile).selectedIndex = 0 
		nFilter = 0 
	End If
	'--------------------------------------
	' LOAD BW PROFILES
	'--------------------------------------
	If vPolicerList(nFilter) <> "" Then g_objIE.Document.All(html_BW_Params)(0).Value = vPolicerList(nFilter) Else g_objIE.Document.All(html_BW_Params)(0).Value = "N/A" End If
	If InStr(vCIR(nFilter),"/") and vCIR(nFilter)  <> "N/A" Then 
		g_objIE.Document.All(html_BW_Params)(1).Value = Split(vCIR(nFilter),"/")(0)
		g_objIE.Document.All(html_BW_Params)(2).Value = Split(vCIR(nFilter),"/")(1)		
	Else 
		g_objIE.Document.All(html_BW_Params)(1).Value = "N/A" 
		g_objIE.Document.All(html_BW_Params)(2).Value = "N/A" 		
	End If
	'
	'  Load CBS/PBS value to the form
	'  Load Either Default value or value from original configuration file
	If g_objIE.Document.All("DefBWprofile").Checked Then 
		If InStr(vCBS(nFilter),"/") and vCBS(nFilter) <> "N/A" Then 
			Select Case ValidateCBS(nFilter)
				Case 1  ' CIR=N and EIR=0
						CBS = CBSdef : PBS = CBS
				Case 2  ' CIR=0 and EIR=N
						CBS = CBSo : PBS = int(CBSo) + int(CBSdef)				
				Case 3  ' CIR=N and EIR=N
						CBS = CBSdef : PBS = 2 * CBS				
			End Select
			g_objIE.Document.All(html_BW_Params)(3).Value = CBS 
			g_objIE.Document.All(html_BW_Params)(4).Value = PBS	
		Else 
			g_objIE.Document.All(html_BW_Params)(3).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(4).Value = "N/A" 		
		End If	
	Else 
		If InStr(vCBS(nFilter),"/") and vCBS(nFilter) <> "N/A" Then 
			g_objIE.Document.All(html_BW_Params)(3).Value = Split(vCBS(nFilter),"/")(0) 
			g_objIE.Document.All(html_BW_Params)(4).Value = Split(vCBS(nFilter),"/")(1) 
		Else 
			g_objIE.Document.All(html_BW_Params)(3).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(4).Value = "N/A" 		
		End If	
		
	End If
	SelectBwProfiles = nFilter
End Function
'------------------------------------------------
'   Function EnableBwProfiles(g_objIE)
'------------------------------------------------
Function EnableBwProfiles(ByRef g_objIE)
	'
	'   Enable all fields by default
	Call DisableIntfButtons(g_obj_IE, false)
	Call DisableBWParams(g_objIE, "BW_Param_1", false)
	Call DisableBWParams(g_objIE, "BW_Param_2", false)
	Call DisableBWParams(g_objIE, "bw_profile_1", false)
	Call DisableBWParams(g_objIE, "bw_profile_2", false)
	Call DisableBWParams(g_objIE, "Policer_BW_Param_1", false)
	Call DisableBWParams(g_objIE, "Policer_BW_Param_2", false)
    '
	'  Disable Bw Profile and Filters if Filter name not configured or unknown
	If g_objIE.document.getElementById("bw_profile_1").length < 2 Then	
		g_objIE.document.getElementById("bw_profile_1").Disabled = True
		For i = 0 to 1
				g_objIE.Document.All("Flt_Intf0" & i )(0).Disabled = True
		Next 
		For i = 0 to 3
				g_objIE.Document.All("Flt_Intf1" & i )(0).Disabled = True
		Next 
	End If 
	If g_objIE.document.getElementById("bw_profile_2").length < 2 Then	
		g_objIE.document.getElementById("bw_profile_2").Disabled = True	
		For i = 0 to 1
				g_objIE.Document.All("Flt_Intf0" & i )(1).Disabled = True
		Next 
		For i = 0 to 3
				g_objIE.Document.All("Flt_Intf1" & i )(1).Disabled = True
		Next 
	End If 
	'
	'   Disable BW Parameters if no Policer found for Filter
	If SelectedOption(g_objIE,"Policer_BW_Param_1") = "N/A" Then	Call DisableBWParams(g_objIE, "Policer_BW_Param_1", True)
	If SelectedOption(g_objIE,"Policer_BW_Param_2") = "N/A" Then	Call DisableBWParams(g_objIE, "Policer_BW_Param_2", True)
	'
	'   Disable BW Parameters if correspondent Left UNI not configured
	If g_objIE.Document.All("UNI_BUTTON00").Value = "N/A" Then 
		Call DisableBWParams(g_objIE, "bw_profile_1", True)
		Call DisableBWParams(g_objIE, "BW_Param_1", True)
		Call DisableBWParams(g_objIE, "Policer_BW_Param_1", True)		
	End If
	If g_objIE.Document.All("UNI_BUTTON01").Value = "N/A" Then 
		Call DisableBWParams(g_objIE, "bw_profile_2", True)
		Call DisableBWParams(g_objIE, "BW_Param_2", True)
		Call DisableBWParams(g_objIE, "Policer_BW_Param_2", True)		
	End If
	'
	'  Disable interface radio button if correspondent UNI/ENNI is not configured 
	For i = 0 to 1
	    If g_objIE.Document.All("UNI_BUTTON0" & i).Value = "N/A" Then 
			g_objIE.Document.All("Flt_Intf0" & i )(0).Disabled = True
			g_objIE.Document.All("Flt_Intf0" & i )(1).Disabled = True
		End If
	Next 
	For i = 0 to 3
	    If g_objIE.Document.All("UNI_BUTTON1" & i).Value = "N/A" Then 
			g_objIE.Document.All("Flt_Intf1" & i )(0).Disabled = True
			g_objIE.Document.All("Flt_Intf1" & i )(1).Disabled = True
		End If
	Next 
End Function
'-------------------------------------------------------
'   Function ApplyFilterList()
'-------------------------------------------------------
Function ApplyFilterList(ByRef g_objIE, ByRef FilterCurrent1, ByRef FilterCurrent2)
Dim nFilter, strFilter
    ApplyFilterList = ""
	'
	'   Filter 1
	If Not g_objIE.document.getElementById("bw_profile_1").Disabled and SelectedOption(g_objIE,"Policer_BW_Param_1") <> "N/A" Then 
		nFilter = g_objIE.document.getElementById("bw_profile_1").selectedIndex
		strFilter = g_objIE.document.getElementById("bw_profile_1").Options(nFilter).text
		FilterCurrent1 = strFilter
		' 	
		'   Filter 1  on Node L 
		For i = 0 to 1
			If g_objIE.Document.All("Flt_Intf0" & i )(0).Checked = True Then 
			    '
				'  Create a filter record to be applied: "F:<filtername>:<node name>:<interface name>"
				ApplyFilterList = ApplyFilterList & " /ARG F /ARG " & "L" & ":" & strFilter & ":" & i
			End If 
		Next 
		' 	
		'   Filter 1 on Node R 
		For i = 0 to 3
			If g_objIE.Document.All("Flt_Intf1" & i )(0).Checked = True Then 
			    '
				'  Create a filter record to be applied: "F:<filtername>:<node name>:<interface name>"
				ApplyFilterList = ApplyFilterList & " /ARG F /ARG " & "R" & ":" & strFilter & ":" & i
			End If 
		Next 
	End If
	Call TrDebug ("ApplyFilter List 1: " & ApplyFilterList, "" , objDebug, MAX_WIDTH, 1, 1)
	'
	'   Filter 2
	If Not g_objIE.document.getElementById("bw_profile_2").Disabled and SelectedOption(g_objIE,"Policer_BW_Param_2") <> "N/A" Then 
		nFilter = g_objIE.document.getElementById("bw_profile_2").selectedIndex
		strFilter = g_objIE.document.getElementById("bw_profile_2").Options(nFilter).text
		FilterCurrent2 = strFilter
		' 	
		'   Filter 2  on Node L 
		For i = 0 to 1
			If g_objIE.Document.All("Flt_Intf0" & i )(1).Checked = True Then 
			    '
				'  Create a filter record to be applied: "F:<filtername>:<node name>:<interface name>"
				'If ApplyFilterList <> "" Then ApplyFilterList = ApplyFilterList & ","				
				ApplyFilterList = ApplyFilterList & " /ARG F /ARG " & "L" & ":" & strFilter & ":" & i
			End If 
		Next 
		' 	
		'   Filter 2 on Node R
		For i = 0 to 3
			If g_objIE.Document.All("Flt_Intf1" & i )(1).Checked = True Then 
			    '
				'  Create a filter record to be applied: "F:<filtername>:<node name>:<interface name>"
				'If ApplyFilterList <> "" Then ApplyFilterList = ApplyFilterList & "/ARG -F "
				ApplyFilterList = ApplyFilterList & " /ARG F /ARG " & "R" & ":" & strFilter & ":" & i
			End If 
		Next 
	End If
	Call TrDebug ("ApplyFilter List all: " & ApplyFilterList, "" , objDebug, MAX_WIDTH, 1, 1)	
End Function
'----------------------------------------------
'  Function CopyFileToClipboard()
'-----------------------------------------------
Function CopyFileToClipboard(strFileName)
Dim oExec, oIn, g_objShell,ClipBoardText,objDataFileName
    Set g_objShell = WScript.CreateObject("WScript.Shell")
	Set oExec = g_objShell.Exec("clip")
	Set oIn = oExec.stdIn
	On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.OpenTextFile(strFileName,1,True)
		ClipBoardText = objDataFileName.ReadAll
		objDataFileName.close
		oIn.WriteLine ClipBoardText
		oIn.Close
		If Err.Number > 0 Then MsgBox "Something went wrong!" & chr(13) & "Can't copy file content to clipboard." & chr(13) & strFileName
		Err.clear
	On Error Goto 0
End Function
'-------------------------------------------------
'  Function MainMenuButton()
'-------------------------------------------------
Function MainMenuButton(ButtonID, ButtonName,CellHight,CellWidth,ButtonX,ButtonY,BGColor,TXTColor,ColSpan,ButtonAlign,ButtonText,EventID, nFontH, css_class)
Dim nFontSize_12
	nFontSize_12 = 12
	If css_class <> "" Then 
		MainMenuButton = "<td colspan=""" & ColSpan & """ style=""border-style: none; background-color: Transparent;""align=""" & ButtonAlign &_
			""" class=""oa1"" height=""" & CellHight & """ width=""" & CellWidth & """ >" & _
			"<button class=""" & css_class & """ style=' width:" & ButtonX & "px; height:" & ButtonY & "px;' " &_
			"id='Button"& ButtonID &"' name='"& ButtonName &"'" &_ 
			"onclick=document.all('ButtonHandler').value='"& EventID &"'" 
	Else 
		MainMenuButton = "<td colspan=""" & ColSpan & """ style=""border-style: none; background-color: Transparent;""align=""" & ButtonAlign &_
			""" class=""oa1"" height=""" & CellHight & """ width=""" & CellWidth & """ >" & _
			"<button style='font-weight: bold; border-style: None; background-color: " & BGColor & "; color: " & TXTColor & "; width:" &_
			ButtonX & ";height:" & ButtonY & ";font-size: " & nFontH & "pt;' " &_
			"id='Button"& ButtonID &"' name='"& ButtonName &"'" &_ 
			"onclick=document.all('ButtonHandler').value='"& EventID &"'" 
	End If
		MainMenuButton = MainMenuButton & " >"& ButtonText &"</button>" & "</td>"
End Function
'-------------------------------------------------------
'   Function ApplyBwProfile()
'-------------------------------------------------------
Function ApplyBwProfile(ByRef g_objIE)
Dim nFilter, strFilter, CIR, PIR, CBS, EBS
    ApplyBwProfile = ""
	'
	'   Profile 1
	If SelectedOption(g_objIE, "Policer_BW_Param_1") <> "N/A" Then 
		CIR = g_objIE.document.All("BW_Param_1")(1).Value
		PIR = g_objIE.document.All("BW_Param_1")(2).Value
		CBS = g_objIE.document.All("BW_Param_1")(3).Value
		EBS = g_objIE.document.All("BW_Param_1")(4).Value
		strFilter = SelectedOption(g_objIE, "Policer_BW_Param_1")
				
		ApplyBwProfile = ApplyBwProfile & " /ARG P /ARG " & "L" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
		ApplyBwProfile = ApplyBwProfile & " /ARG P /ARG " & "R" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
	End If
	Call TrDebug ("ApplyFilter List 1: " & ApplyBwProfile, "" , objDebug, MAX_WIDTH, 1, 1)
	'
	'   Profile 2
	If SelectedOption(g_objIE, "Policer_BW_Param_2") <> "N/A" Then 
		CIR = g_objIE.document.All("BW_Param_2")(1).Value
		PIR = g_objIE.document.All("BW_Param_2")(2).Value
		CBS = g_objIE.document.All("BW_Param_2")(3).Value
		EBS = g_objIE.document.All("BW_Param_2")(4).Value
		strFilter = SelectedOption(g_objIE, "Policer_BW_Param_2")
				
		ApplyBwProfile = ApplyBwProfile & " /ARG P /ARG " & "L" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
		ApplyBwProfile = ApplyBwProfile & " /ARG P /ARG " & "R" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
	End If
	Call TrDebug ("ApplyFilter List 1: " & ApplyBwProfile, "" , objDebug, MAX_WIDTH, 1, 1)
	Call TrDebug ("ApplyFilter List all: " & ApplyBwProfile, "" , objDebug, MAX_WIDTH, 1, 1)	
End Function
'-------------------------------------------------------
'   Function ValidateBWProfile(objIE)
'-------------------------------------------------------
Function ValidateBWProfile(ByRef g_objIE) 
ValidateBWProfile = True

End Function
'------------------------------------------------
'   Function DisableBWParams(g_objIE, html_BW_Params, bDisable)
'------------------------------------------------
Function DisableBWParams(ByRef g_objIE, html_BW_Params, bDisabled)
Dim BW_Param, Item
    set BW_Param = g_objIE.document.All(html_BW_Params)
	For Each Item in BW_Param
		Item.disabled = bDisabled
	Next 
	set BW_Param = Nothing
End Function
'------------------------------------------------
'   Function DisableIntfButtons(g_objIE, bDisable)
'------------------------------------------------
Function DisableIntfButtons(g_objIE,bDisable)
	On Error Resume Next
		For j = 0 to 1
			For i = 0 to 3
				g_objIE.Document.All("Flt_Intf" & j & i )(0).Disabled = bDisable
				g_objIE.Document.All("Flt_Intf" & j & i )(1).Disabled = bDisable				
			Next
		Next
	On Error Goto 0
End Function
'------------------------------------------------
'   Function ValidateCBS(nFilter)
'------------------------------------------------
Function ValidateCBS(nFilter)
Dim CIR, PIR, CBS, PBS
	If vCIR(nFilter) = "" Then ValidateCBS = 0 : Exit Function : End If
	CIR = Trim(Split(vCIR(nFilter),"/")(0))
	PIR = Trim(Split(vCIR(nFilter),"/")(1))
	CBS = Trim(Split(vCBS(nFilter),"/")(0))
	PBS = Trim(Split(vCBS(nFilter),"/")(1))
	if InStr(CIR,PIR) and InStr(PIR,CIR) Then ValidateCBS = 1 : Exit Function : End If
	if CIR = CIRo or CBS = CBSo Then ValidateCBS = 2 : Exit Function : End If		
	if CIR <> PIR 			    Then ValidateCBS = 3 : Exit Function : End If		
End Function
'-----------------------------------------------------------------------
'   GetFilterList_New(vConfigFileLeft, ByRef vFilterList, ByRef vFilterProfileList, ByRef vPolicerList, ByRef vCIR, ByRef vCBS, nDebug)
'-----------------------------------------------------------------------
Function GetFilterList_New(vConfigFileLeft, ByRef vFilterList, ByRef vFilterProfileList, ByRef vPolicerList, ByRef vCIR, ByRef vCBS, nDebug)
Dim nFilter, nPolicer,FW_Start, FoundFilter, FoundPolicer, strLine, nBraces
	Redim vFilterList(1)
	Redim vPolicerList(1)
	Redim vFilterProfileList(1)
	Redim vCIR(1)
	Redim vCBS(1)
	vFilterList(0) = "N/A"
	vPolicerList(0) = "N/A"
	vFilterProfileList(0) = "N/A"
	GetFilterList_New = 0
	nFilter = 0
	nPolicer = 0
	nBraces = 10000
	FoundFilter = False
	FoundPolicer = False
	For Each strLine in vConfigFileLeft
		strLine = Trim(strLine)
        'If nBraces => 0 and nBraces < 10 Then Call TrDebug ("GetFilterList: " & strLine, nBraces,objDebug, MAX_LEN, 1, nDebug)		
		Select Case nBraces
			Case 1
				If FoundPolicer Then nPolicer = nPolicer + 1
				FoundPolicer = False
				If Instr(strLine,"policer") Then 
					FoundPolicer = True
					Redim Preserve vPolicerList(nPolicer + 1)
					Redim Preserve vCIR(nPolicer + 1)
					Redim Preserve vCBS(nPolicer + 1)
					vPolicerList(nPolicer) = Split(strLine," ")(1)
					Call TrDebug ("GetFilterList: FOUND POLICER:", vPolicerList(nPolicer),objDebug, MAX_LEN, 1, nDebug)
				End If
			Case 2
			    If FoundFilter Then nFilter = nFilter + 1
				FoundFilter = False
				If Instr(strLine,"filter ") Then 
					FoundFilter = True				
					Redim Preserve vFilterList(nFilter + 1)
					Redim Preserve vFilterProfileList(nFilter + 1)
					vFilterProfileList(nFilter) = "N/A"
					vFilterList(nFilter) = Split(strLine," ")(1)
					Call TrDebug ("GetFilterList: FOUND FILTER: " & nFilter, vFilterList(nFilter),objDebug, MAX_LEN, 1, nDebug)
				End If				
			Case 3
					If InStr(strLine,"committed-information-rate") 	Then vCIR(nPolicer) = Split(Split(strLine," ")(1),";")(0)
					If InStr(strLine,"bandwidth-limit") 			Then vCIR(nPolicer) = Split(Split(strLine," ")(1),";")(0)
					If InStr(strLine,"peak-information-rate") 		Then vCIR(nPolicer) = vCIR(nPolicer) & "/" & Split(Split(strLine," ")(1),";")(0)
					If InStr(strLine,"committed-burst-size") 		Then vCBS(nPolicer) = Split(Split(strLine," ")(1),";")(0)
					If InStr(strLine,"burst-size-limit") 			Then vCBS(nPolicer) = Split(Split(strLine," ")(1),";")(0)
					If InStr(strLine,"peak-burst-size") 			Then vCBS(nPolicer) = vCBS(nPolicer) & "/" & Split(Split(strLine,"peak-burst-size ")(1),";")(0)				
					'Call TrDebug ("GetFilterList: FOUND PARAMETER:", strLine,objDebug, MAX_LEN, 1, nDebug)
			Case 5
			    If  InStr(strLine,"policer") and Instr(strLine,"three-color-policer") = 0 Then 
					vFilterProfileList(nFilter) = vFilterProfileList(nFilter) & ":" & Split(Split(strLine," ")(1),";")(0)
					Call TrDebug ("GetFilterList: " & strLine, nBraces,objDebug, MAX_LEN, 1, nDebug)		
					Call TrDebug ("GetFilterList: FOUND POLICER IN FILTER:", vFilterProfileList(nFilter),objDebug, MAX_LEN, 1, nDebug)
				End If
			Case 6
				If Instr(strLine,"two-rate") Then
					vFilterProfileList(nFilter) = vFilterProfileList(nFilter) & ":" & Split(Split(strLine," ")(1),";")(0)
					Call TrDebug ("GetFilterList: " & strLine, nBraces,objDebug, MAX_LEN, 1, nDebug)		
					Call TrDebug ("GetFilterList: FOUND POLICER IN FILTER: " & nFilter , vFilterProfileList(nFilter),objDebug, MAX_LEN, 1, nDebug)
					
				End If
			Case Else 
				If InStr(strLine,"firewall ") Then 
					FW_Start = True 
					nBraces = 0 
					Call TrDebug ("GetFilterList: FOUND FW:", strLine,objDebug, MAX_LEN, 1, nDebug)
				End If				
		End Select
		If InStr(strLine,"{") Then nBraces = nBraces + 1
		If InStr(strLine,"}") Then nBraces = nBraces - 1
		if nBraces = 0 Then Exit For
	Next
	If UBound(vPolicerList) = 0 Then Redim vPolicerList(1) : vPolicerList(0) = "N/A" : End If
	GetFilterList_New = nFilter
End Function
'------------------------------------------------
'   Function SelectBwProfiles_New(g_objIE, html_BW_Profile, html_BW_Params)
'------------------------------------------------
Function SelectBwProfiles_New(ByRef g_objIE, html_BW_Profile, html_BW_Params, nFilter)
Dim html_Policer_list, vList, strPolicer, nPolicer

    html_Policer_list = "Policer_" & html_BW_Params
	' nFilter = g_objIE.document.getElementById(html_BW_Profile).selectedIndex
	'--------------------------------------
	' LOAD BW PROFILES
	'--------------------------------------
	If UBound(Split(vFilterProfileList(nFilter),":")) = 0 Then 
		g_objIE.document.getElementById(html_Policer_list).length = 1
		g_objIE.document.getElementById(html_Policer_list).Options(0).Text = "N/A"
	Else 
		vList = Split(vFilterProfileList(nFilter),":")
		g_objIE.document.getElementById(html_Policer_list).length = 0
		For n = 1 to UBound(Split(vFilterProfileList(nFilter),":"))
			strPolicer = Split(vFilterProfileList(nFilter),":")(n)
			Select Case strPolicer
				Case "N/A"
				Case "" 
					Exit For
				Case Else
					g_objIE.document.getElementById(html_Policer_list).length = n
					g_objIE.document.getElementById(html_Policer_list).Options(n-1).text = strPolicer
			End Select
		Next 
		g_objIE.document.getElementById(html_Policer_list).selectedIndex = 0
	End If
'	If vPolicerList(nFilter) <> "" Then g_objIE.Document.All(html_BW_Params)(0).Value = vPolicerList(nFilter) Else g_objIE.Document.All(html_BW_Params)(0).Value = "N/A" End If
	strPolicer = g_objIE.document.getElementById(html_Policer_list).Options(0).text
	Select Case strPolicer
		Case "N/A"
			g_objIE.Document.All(html_BW_Params)(1).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(2).Value = "N/A" 		
			g_objIE.Document.All(html_BW_Params)(3).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(4).Value = "N/A" 		
		Case Else 
			For nPolicer = 0 to UBound(vPolicerList) - 1
				if strPolicer = vPolicerList(nPolicer) Then Exit for 
			Next 
			g_objIE.Document.All(html_BW_Params)(1).Value = Split(vCIR(nPolicer),"/")(0)
			g_objIE.Document.All(html_BW_Params)(2).Value = Split(vCIR(nPolicer),"/")(1)		
			If g_objIE.Document.All("DefBWprofile").Checked Then 
				Select Case ValidateCBS(nPolicer)
					Case 1  ' CIR=N and EIR=0
							CBS = CBSdef : PBS = CBS
					Case 2  ' CIR=0 and EIR=N
							CBS = CBSo : PBS = int(CBSo) + int(CBSdef)				
					Case 3  ' CIR=N and EIR=N
							CBS = CBSdef : PBS = 2 * CBS				
				End Select
				g_objIE.Document.All(html_BW_Params)(3).Value = CBS 
				g_objIE.Document.All(html_BW_Params)(4).Value = PBS	
			Else 
				g_objIE.Document.All(html_BW_Params)(3).Value = Split(vCBS(nPolicer),"/")(0) 
				g_objIE.Document.All(html_BW_Params)(4).Value = Split(vCBS(nPolicer),"/")(1) 
			End If
	End Select 
	SelectBwProfiles_New = nPolicer
End Function
'------------------------------------------------
'   Function LoadBwProfiles_New(g_objIE, html_BW_Profile, html_BW_Params,nFilter)
'------------------------------------------------
Function LoadBwProfiles_New(ByRef g_objIE, html_BW_Profile, html_BW_Params,nFilter)
Dim CBS, PBS

	If GetFilterList_New(vConfigFileLeft, vFilterList,vFilterProfileList, vPolicerList, vCIR, vCBS, nDebug)=0 Then 
		g_objIE.document.getElementById(html_BW_Profile).length = 1
		g_objIE.document.getElementById(html_BW_Profile).Options(0).Text = "N/A"
		g_objIE.document.getElementById(html_BW_Profile).SelectedIndex = 0
		g_objIE.document.getElementById(html_BW_Profile).Disabled = True
	Else
		g_objIE.document.getElementById(html_BW_Profile).Disabled = False
		For n = 0 to UBound(vFilterList) - 1
			g_objIE.document.getElementById(html_BW_Profile).length = n + 1
			g_objIE.document.getElementById(html_BW_Profile).Options(n).text = vFilterList(n) 
		Next
		If Int(nFilter) <= n Then g_objIE.document.getElementById(html_BW_Profile).selectedIndex = nFilter
	End If
	Call SelectBwProfiles_New(g_objIE, html_BW_Profile, html_BW_Params, nFilter)
	For i = 0 to 1
		g_objIE.Document.All("Flt_Intf0" & i )(0).Disabled = False
		g_objIE.Document.All("Flt_Intf0" & i )(1).Disabled = False
		g_objIE.Document.All("Flt_Intf0" & i )(0).Checked = False
		g_objIE.Document.All("Flt_Intf0" & i )(1).Checked = False
	Next 
	For i = 0 to 3
	    g_objIE.Document.All("Flt_Intf1" & i )(0).Disabled = False
		g_objIE.Document.All("Flt_Intf1" & i )(1).Disabled = False
	    g_objIE.Document.All("Flt_Intf1" & i )(0).Checked = False
		g_objIE.Document.All("Flt_Intf1" & i )(1).Checked = False
	Next 	
End Function
'------------------------------------------------
'   Function SelectPolicer_New(g_objIE, html_BW_Params)
'------------------------------------------------
Function SelectPolicer_New(ByRef g_objIE, html_BW_Params)
Dim vList, strPolicer, nPolicer, html_Policer_list
    html_Policer_list = "Policer_" & html_BW_Params
	strPolicer = SelectedOption(g_objIE, html_Policer_list)
	Select Case strPolicer
		Case "N/A"
			g_objIE.Document.All(html_BW_Params)(1).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(2).Value = "N/A" 		
			g_objIE.Document.All(html_BW_Params)(3).Value = "N/A" 
			g_objIE.Document.All(html_BW_Params)(4).Value = "N/A" 		
		Case Else 
			For nPolicer = 0 to UBound(vPolicerList) - 1
				if strPolicer = vPolicerList(nPolicer) Then Exit for 
			Next 
			g_objIE.Document.All(html_BW_Params)(1).Value = Split(vCIR(nPolicer),"/")(0)
			g_objIE.Document.All(html_BW_Params)(2).Value = Split(vCIR(nPolicer),"/")(1)		
			If g_objIE.Document.All("DefBWprofile").Checked Then 
				Select Case ValidateCBS(nPolicer)
					Case 1  ' CIR=N and EIR=0
							CBS = CBSdef : PBS = CBS
					Case 2  ' CIR=0 and EIR=N
							CBS = CBSo : PBS = int(CBSo) + int(CBSdef)				
					Case 3  ' CIR=N and EIR=N
							CBS = CBSdef : PBS = 2 * CBS				
				End Select
				g_objIE.Document.All(html_BW_Params)(3).Value = CBS 
				g_objIE.Document.All(html_BW_Params)(4).Value = PBS	
			Else 
				g_objIE.Document.All(html_BW_Params)(3).Value = Split(vCBS(nPolicer),"/")(0) 
				g_objIE.Document.All(html_BW_Params)(4).Value = Split(vCBS(nPolicer),"/")(1) 
			End If
	End Select 
	SelectPolicer_New = nPolicer
End Function
'-------------------------------------------------------------
'   Function SelectedOption(ByRef g_objIE, htmlID)
'-------------------------------------------------------------
Function SelectedOption(ByRef g_objIE, htmlID)
Dim nIndex 
	nIndex = g_objIE.document.getElementById(htmlID).SelectedIndex
	SelectedOption = g_objIE.document.getElementById(htmlID).Options(nIndex).Text
End Function 
'-------------------------------------------------------
'   Function ApplyBwProfile_All()
'-------------------------------------------------------
Function ApplyBwProfile_All(ByRef g_objIE)
Dim nFilter, strFilter, CIR, PIR, CBS, EBS, nPolicer, nIndex
    ApplyBwProfile_All = ""
	'
	'   Profile 1
	If g_objIE.document.getElementById("Policer_BW_Param_1").Options(0).Text <> "N/A" Then 
		For nIndex = 0 to g_objIE.document.getElementById("Policer_BW_Param_1").length - 1			
			For nPolicer = 0 to Ubound(vPolicerList) - 1
				if vPolicerList(nPolicer) = g_objIE.document.getElementById("Policer_BW_Param_1").Options(nIndex).Text Then Exit For
			Next 
			CIR = Split(vCIR(nPolicer),"/")(0)
			PIR = Split(vCIR(nPolicer),"/")(1)
			If g_objIE.Document.All("DefBWprofile").checked Then 
				Select Case ValidateCBS(nPolicer)
					Case 1  ' CIR=N and EIR=0
							CBS = CBSdef : EBS = CBS
					Case 2  ' CIR=0 and EIR=N
							CBS = CBSo : EBS = int(CBSo) + int(CBSdef)				
					Case 3  ' CIR=N and EIR=N
							CBS = CBSdef : EBS = 2 * CBS				
				End Select				
			Else 
				CBS = Split(vCBS(nPolicer),"/")(0)
				EBS = Split(vCBS(nPolicer),"/")(1)
			End If
			strFilter = vPolicerList(nPolicer)
			ApplyBwProfile_All = ApplyBwProfile_All & " /ARG P /ARG " & "L" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
			ApplyBwProfile_All = ApplyBwProfile_All & " /ARG P /ARG " & "R" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
		Next
	End If
	Call TrDebug ("ApplyFilter List 1: " & ApplyBwProfile_All, "" , objDebug, MAX_WIDTH, 1, 1)
	'
	'   Profile 2
	If g_objIE.document.getElementById("Policer_BW_Param_2").Options(0).Text <> "N/A" Then 
		For nIndex = 0 to g_objIE.document.getElementById("Policer_BW_Param_2").length - 1
			For nPolicer = 0 to Ubound(vPolicerList) - 1
				if vPolicerList(nPolicer) = g_objIE.document.getElementById("Policer_BW_Param_2").Options(nIndex).Text Then Exit For
			Next 
			CIR = Split(vCIR(nPolicer),"/")(0)
			PIR = Split(vCIR(nPolicer),"/")(1)
			If g_objIE.Document.All("DefBWprofile").checked Then 
				Select Case ValidateCBS(nPolicer)
					Case 1  ' CIR=N and EIR=0
							CBS = CBSdef : EBS = CBS
					Case 2  ' CIR=0 and EIR=N
							CBS = CBSo : EBS = int(CBSo) + int(CBSdef)				
					Case 3  ' CIR=N and EIR=N
							CBS = CBSdef : EBS = 2 * CBS				
				End Select				
			Else 
				CBS = Split(vCBS(nPolicer),"/")(0)
				EBS = Split(vCBS(nPolicer),"/")(1)
			End If
			strFilter = vPolicerList(nPolicer)
			ApplyBwProfile_All = ApplyBwProfile_All & " /ARG P /ARG " & "L" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
			ApplyBwProfile_All = ApplyBwProfile_All & " /ARG P /ARG " & "R" & ":" & strFilter & ":" & CIR & ":" & PIR & ":" & CBS & ":" & EBS 
		Next
	End If
	Call TrDebug ("ApplyFilter List all: " & ApplyBwProfile_All, "" , objDebug, MAX_WIDTH, 1, 1)	
End Function
'-------------------------------------
'  Function BrowseForFile()
'-------------------------------------
'Function BrowseForFile(strFolder)
'    Dim shell : Set shell = CreateObject("Shell.Application")
'	On Error Resume Next
'    Dim file : Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, strFolder)
'	MsgBox file
'    BrowseForFile = file.self.Path
'		MsgBox BrowseForFile
'	On Error Goto 0
	' Set shell = Nothing
'End Function
Function BrowseForFile()
    With CreateObject("WScript.Shell")
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
        Dim tempName : tempName = fso.GetTempName() & ".hta"
        Dim path : path = "HKCU\Volatile Environment\MsgResp"
        With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
            "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
            "close();</script>"
            .Close
        End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
    End With
End Function
'------------------------------------------------------------------
'   Function OpenLogSession(ByRef objDebug, ByRef strDebugFile, bMultipleInstanceAllowed, bShowLog, bVerbose)
'------------------------------------------------------------------
Function OpenLogSession(ByRef objDebug, ByRef strDebugFile, UtilsFolder, bMultipleInstanceAllowed, bShowLog, bVerbose)
Dim nIndex, strErrorLog,strNewInstanceLog, objError, INSTANCE_LOG, nError, DEBUG_FILE, MyScriptName
Dim my_objShell
	Set my_objShell = CreateObject("WScript.Shell")
	MyScriptName = Split(Wscript.ScriptFullName,"\")(UBound(Split(Wscript.ScriptFullName,"\")))
    nError = 0
	nLenEnd = InStrRev(strDebugFile,"\")
	strErrorLog = Left(strDebugFile,nLenEnd) & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0) & "_" & My_Random(1,9999) & "_Error.log"
	If InStr(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".") Then 
	   INSTANCE_LOG = Left(strDebugFile,nLenEnd) & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0) & "-inst-<index>." & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(1)
	Else 
	   INSTANCE_LOG = Left(strDebugFile,nLenEnd) & Right(strDebugFile,Len(strDebugFile) - nLenEnd) & "-inst-<index>"
	End If   
	nIndex = 0
    Set objError = objFSO.OpenTextFile(strErrorLog,ForAppending,True)
	Do
		On Error Resume Next
		Err.Clear
		Set objDebug = objFSO.OpenTextFile(strDebugFile,ForAppending,True)
		Select Case Err.Number
			Case 0
				Exit Do
			Case 70
				nIndex = nIndex + 1
				Select Case nIndex
                   	Case 1
                       	If bMultipleInstanceAllowed Then 
					        strDebugFile = Replace(INSTANCE_LOG,"<index>",nIndex)
						Else 
							If bVerbose Then MsgBox MyScriptName & ": Another instance of the script is already running." & chr(13) & "Exit now"
							objError.WriteLine Date() & " " & Time() & ": ERROR:  Another instances of the script is already running. Exit now"
                            nError = 1
							Exit Do
						End if 
                    Case 2					
                           strDebugFile = Replace(INSTANCE_LOG,"<index>",nIndex)
					Case 3
					    If bVerbose Then MsgBox MyScriptName & ": Two other instances of the script are already running. " & chr(13) & "Exit now"
						objError.WriteLine Date() & " " & Time() & ": ERROR:  Two other instances of the script are already running. Exit now"
						nError = 3
						Exit Do
				End Select
				wscript.sleep 500
			Case Else 
			    If bVerbose Then MsgBox MyScriptName & ": Can't open log file" & chr(13) & "Error: #" & Err.Number & ": " &  Err.Description 
			    objError.WriteLine Date() & " " & Time() & "ERROR: Can't open log file" 
				objError.WriteLine Date() & " " & Time() & "Error: #" & Err.Number & ": " &  Err.Description 
			    nError = 1000
				Exit Do
		End Select
	Loop
	On Error goto 0
    If nError > 0 Then 
	   OpenLogSession = False
	   If IsObject(objError) Then objError.Close : End If
	   If bShowLog Then my_objShell.Run "notepad.exe " & strErrorLog,1
	   set my_objShell = Nothing
	   Exit Function
	End If 
	' Open tail -f to stream log mesages into desktop
	'wscript.sleep 1000
	If bShowLog Then 
		strLaunch = UtilsFolder & "\tail.exe -n 10 -f " & """" & strDebugFile & """"
		DEBUG_FILE = Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0)
		If Not GetWinAppPID(strPID, strParrentID, DEBUG_FILE, "tail.exe", nDebug) Then 
			my_objShell.run (strLaunch)
		Else
			Call FocusToParentWindow(strPID)
		End If
	End If
	' Exit Function
	Set my_objShell = Nothing
    OpenLogSession = True
    If IsObject(objError) Then objError.Close : End If
End Function 
'-----------------------------------
'
'-----------------------------------
Function ValidateFS(nDebug) 
Dim objFile, vFileLines, strLine, strFile
    If Not objFSO.FolderExists(strDirectoryScript) Then 
	    MsgBox "Can't find folder " & chr(13) & strDirectoryScript
		ValidateFS = False
		Exit Function
	End If 
    If Not objFSO.FolderExists(strDirectoryBin) Then 
	    MsgBox "Can't find folder " & chr(13) & strDirectoryBin 
		ValidateFS = False
		Exit Function
	End If 
    If Not objFSO.FolderExists(strDirectoryLog) Then 
	    objFSO.CreateFolder(strDirectoryLog)
	End If 
    If Not objFSO.FolderExists(strDirectoryTmp) Then 
	    objFSO.CreateFolder(strDirectoryTmp)
	End If 
    If Not objFSO.FileExists(strFileSession) Then 
	    Set objFile = objFSO.OpenTextFile(strFileSession,2,True)
		objFile.writeline "Settings file = " & strFileSettings
		objFile.close
		Set objFile = Nothing
	Else 
	    ' Get names of the topology and settings file 
	    Call GetFileLineCountSelect(strFileSession, vFileLines,"", "", "", nDebug)
		strFile = ""
		For Each strLine in vFileLines
		    If strLine = "" Then Exit For
		    Select case Trim(Split(strLine, "=")(0)) 
				Case "Settings file"
					strFile = Trim(Split(strLine, "=")(1))
		    End Select 
		Next
		' If Settings file name not found then use default value		
		If strFile <> "" Then 
			strFileSettings = strFile
		Else 
			Set objFile = objFSO.OpenTextFile(strFileSession,2,True)
			objFile.writeline "Settings file = " & strFileSettings
			objFile.close
			Set objFile = Nothing
		End If 
	End If 
    ValidateFS = True	
End Function 
'-------------------------------------------
' Function VaildateFTP(ByRef strFTP_Folder)
'-------------------------------------------
Function VaildateFTP(ByRef strFTP_Folder)
	On Error Resume Next
		Err.Clear
		VaildateFTP = True
		strFTP_Folder = objShell.RegRead(FTP_REG_INSTALL)
		if Err.Number <> 0 Then 
			vvMsg(0,0) = "WARNING: CAN'T FIND FileZilla Server Folder"	           : vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Make sure that FileZilla Server Installed correctly"     : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1			
			vvMsg(2,0) = "If You are using other FTP Server " & _
                		 "make sure that FTP Login/Password are " & _
						 "the same as under MEF Loader Settings. " & _
						 "Folder with DUT Configuration files should " & _
						 "be used as FTP User Homedirectory"                      	: vvMsg(2,1) = "normal" : vvMsg(1,2) = HttpTextColor1						
			vvMsg(3,0) = "" : vvMsg(4,0) = "" :vvMsg(5,0) = "" :vvMsg(6,0) = "" :vvMsg(7,0) = "" :
			Call IE_MSG(vIE_Scale, "Error",vvMsg,8,Null)
			VaildateFTP = False
			strFTP_Folder = "C:"
		End If
		If Right(strFTP_Folder,1) = "\" Then strFTP_Folder = Left(strFTP_Folder,Len(strFTP_Folder)-1)
	On Error Goto 0
End Function
'--------------------------------------------
'  Function ValidateSecureCRT(strCRT_SessionFolder,strCRT_InstallFolder)
'--------------------------------------------
Function ValidateSecureCRT(strCRT_SessionFolder,strCRT_InstallFolder)
	On Error Resume Next
	    ValidateSecureCRT = True
		Err.Clear
		strCRT_InstallFolder = objShell.RegRead(CRT_REG_INSTALL)
		if Err.Number <> 0 Then 
			vvMsg(0,0) = "WARNING: CAN'T FIND SecureCRT Install Folder"	: vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Make sure that Secure CRT Application was " & _
			              "installed on your system correctly"	        : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2			
			vvMsg(2,0) = ""         									: vvMsg(2,1) = "bold" : vvMsg(2,2) = HttpTextColor1
			Call IE_MSG(vIE_Scale, "Error",vvMsg,2,Null)
            ValidateSecureCRT = False
			strCRT_InstallFolder = "C:"
		End If
		If Right(strCRT_InstallFolder,1) = "\" Then strCRT_InstallFolder = Left(strCRT_InstallFolder,Len(strCRT_InstallFolder)-1)
		strCRT_SessionFolder = objShell.RegRead(CRT_REG_SESSION)
		if Err.Number <> 0 Then 
			vvMsg(0,0) = "WARNING: CAN'T FIND SecureCRT Session Folder"	              : vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Once SecureCRT application was installed run it, " & _
			             "so that default session configuration will be created"	  : vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
			vvMsg(2,0) = "Exit . . ."									              : vvMsg(2,1) = "bold" : vvMsg(1,2) = HttpTextColor1
			Call IE_MSG(vIE_Scale, "Error",vvMsg,3,Null)
            ValidateSecureCRT = False
			strCRT_SessionFolder = "C:"
		End If
		If Right(strCRT_SessionFolder,1) = "\" Then strCRT_SessionFolder = Left(strCRT_SessionFolder,Len(strCRT_SessionFolder)-1)
		strCRT_SessionFolder = strCRT_SessionFolder & "\Sessions"
	On Error Goto 0
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
'---------------------------------------------------
'  Sub LoadScriptsNames()
'---------------------------------------------------
Sub LoadScriptsNames()
Dim nInventory, nIndex,vLines
	nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Scripts","","",0)
	For nIndex = 0 to nInventory - 1
		Select Case Split(vLines(nIndex),"=")(0)
			Case "UPLOAD"
				VBScript_Upload_Config = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "DOWNLOAD"
			    VBScript_DNLD_Config = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "FWFILTER"
			    VBScript_FWF_Apply = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "FTPUSER"
			    VBScript_FTP_User = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "SETNODE"
				VBScript_Set_Node = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "BWPROFILE"
			    VBScript_BWP_Apply = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
			Case "FLT_BWP_ROFILE"
				VBScript_FLT_and_BWP_Apply = strDirectoryWork & "\" & Split(vLines(nIndex),"=")(1)
		End Select
	Next
End Sub
'---------------------------------------------------
'  Sub GetCfgLoaderVersion()
'---------------------------------------------------
Sub GetCfgLoaderVersion()
Dim Inventory, nIndex, vLines
	nInventory = GetFileLineCountByGroup(strFileSettings, vLines,"Version","","",0)
	For nIndex = 0 to nInventory - 1
		Select Case Split(vLines(nIndex),"=")(0)
			Case "VERSION"
				strVersion = Split(vLines(nIndex),"=")(1)
		End Select
	Next
End Sub
'---------------------------------------------------
'  Sub GetXLStemplates()
'---------------------------------------------------
Sub GetXLStemplates()
Dim nCount, nInd, nInd1, vLines, n
	nCount = GetFileLineCountByGroup(strFileSettings, vLines,"Templates","","",1)
	nInd = 0 : nInd1 = 0
	For n = 0 to nCount - 1
		Select Case Split(vLines(n),"=")(0)
			Case Template
				vTemplates(nInd) = Split(vLines(n),"=")(1) : nInd = nInd + 1
				Call TrDebug("GetTemplates: ", vTemplates(nInd-1), objDebug, MAX_LEN, 1, 1)
			Case Orig_Folder
			    strTempOrigFolder = Split(vLines(n),"=")(1)
				vSettings(8) = vLines(n)
				Call TrDebug("GetTemplates: ", strTempOrigFolder, objDebug, MAX_LEN, 1, 1)
			Case Dest_Folder
			    strTempDestFolder = Split(vLines(n),"=")(1)
				vSettings(9) = vLines(n)
                Call TrDebug("GetTemplates: ", strTempDestFolder, objDebug, MAX_LEN, 1, 1)				
			Case WorkBookPrefix
			    vXLSheetPrefix(nInd1) = Split(vLines(n),"=")(1) : nInd1 = nInd1 + 1
				Call TrDebug("GetTemplates: ", vXLSheetPrefix(nInd1-1), objDebug, MAX_LEN, 1, 1)				
		End Select
	Next
End Sub
'------------------------------------------------
'  Sub CreateFoldersForTestedConfigurations(strDirectoryConfig)
'------------------------------------------------
Sub CreateFoldersForTestedConfigurations(strDirectoryConfig)
	strDirectoryBackUp = Replace(strDirectoryBackUp,"<ConfigDir>",strDirectoryConfig)
	strDirectoryTested = Replace(strDirectoryTested,"<ConfigDir>",strDirectoryConfig)


	If Not objFSO.FolderExists(strDirectoryBackUp) Then 
		objFSO.CreateFolder(strDirectoryBackUp) 
	End If
	For n = 0 to Ubound(vSvc) - 1
		If Not objFSO.FolderExists(strDirectoryBackup & "\" & vSvc(n,1)) Then 
			objFSO.CreateFolder(strDirectoryBackup & "\" & vSvc(n,1)) 
		End If
	Next
	If Not objFSO.FolderExists(strDirectoryTested) Then 
		objFSO.CreateFolder(strDirectoryTested) 
	End If
	For n = 0 to Ubound(vSvc) - 1
		If Not objFSO.FolderExists(strDirectoryTested & "\" & vSvc(n,1)) Then 
			objFSO.CreateFolder(strDirectoryTested & "\" & vSvc(n,1)) 
		End If
	Next
End Sub
'------------------------------------------------
'  Sub LoadLastTestedSeries(strFileSessionTmp)
'------------------------------------------------
Sub LoadLastTestedSeries(strFileSessionTmp)
Dim nSessionTmp
	If objFSO.FileExists(strFileSessionTmp) Then 
		nSessionTmp = GetFileLineCount(strFileSessionTmp, vSessionTmp,0)
		strServiceID = vSessionTmp(0)
		strFlavorID = vSessionTmp(1)
		strTaskID = vSessionTmp(2)
	Else
		strServiceID = 0
		strFlavorID = 0
		strTaskID = 0
		vSessionTmp(0) = 0
		vSessionTmp(1) = 0
		vSessionTmp(2) = 0
	End If
	Call TrDebug("GetTestSeries: LAST LOADED SESSION: " & strServiceID & "-" & strFlavorID & "-" & strTaskID, "", objDebug, MAX_LEN, 3, 1)
End Sub
'------------------------------------------------------------------------
'	Function CopyFileToString(strSourceFile)
'------------------------------------------------------------------------
Function CopyFileToString(strSourceFile)
Dim strFileString
Dim objFSO,objSourceFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Const ForAppending = 8
	Const ForWriting = 2
	Const ForReading = 1
	If objFSO.FileExists(strSourceFile) Then 	
		On Error Resume Next
			Err.Clear
			Set objSourceFile = objFSO.OpenTextFile(strSourceFile,ForReading,True)
			Select Case Err.Number
				Case 0 ' Do Nothing
				Case Else 
					Err.Clear
					CopyFileToString = "-1"
					Set objFSO = Nothing
					On Error Goto 0
					Exit Function
			End Select	
        Err.Clear
		strFileString = objSourceFile.ReadAll
		objSourceFile.close		
		Set objSourceFile = Nothing
		If Err.Number > 0 Then
					Err.Clear
					CopyFileToString = "-2"
					Set objFSO = Nothing
					On Error Goto 0
					Exit Function		
		End If 
		On Error Goto 0	
    End If		
	Set objFSO = Nothing
	CopyFileToString = strFileString
End Function 
'-------------------------------------------------
'   Function CreateMessageDiv(WindowH,nTab)
'   Add Hidden Message board to the form
'-------------------------------------------------
Function CreateMessageDiv(WindowH,nTab)
Dim WaitIcon
	WaitIcon = strDirectoryWork & "\bin\loader.gif"
	'
	'  Shadow Div to cover parent window
	CreateMessageDiv = "<div id='divShadow' name='divShadow' " &_
						"style=""position: absolute; top: 0px; left: " & WindowH + nTab & "px;"">" &_
						"</div>" 
	CreateMessageDiv =  CreateMessageDiv & "<div id='divTransparent' name='divTransparent' " &_
						"style=""position: absolute; top: 0px; left: " & WindowH + nTab & "px;"" " &_
						"onclick=document.all('ButtonHandler').value='CLICK_ON_SHADOW';>" &_
							"<img id='wait_icon' src=""" & WaitIcon & """ style=""width: 15%; hight: 15%; "" hidden >" &_							
						"</div>" 
	'
	'  Message board
    CreateMessageDiv = CreateMessageDiv & "<div id='divMessage' name='divMessage'" &_
						"style=""position: absolute; top: 0px; left: " & WindowH + nTab & "px;"">" &_
					"</div>" 
End Function
'---------------------------------------------------------
'   Function FormReload(g_objIE, ByRef vReload, bReload)
'---------------------------------------------------------
Function FormReload(ByRef g_objIE, ByRef vReload, bReload)
	vReload(0) = bReload
	vReload(1) = 0       '  ScrollTop
	On Error Resume Next
	Do
		If Not bReload Then Exit Do
		If Not IsObject(g_objIE) Then Exit Do
		If Not IsObject(g_objIE.document.GetElementById("MainDiv")) Then Exit Do
		vReload(1) = g_objIE.document.GetElementById("MainDiv").scrollTop
		Exit Do 
	Loop
	On Error Goto 0
End Function
'--------------------------------------------------------
'   Function SettingsMenu(TitleH,TitleW,MenuH)
'--------------------------------------------------------
Function SettingsMenu(TitleH,MenuW,MenuH,nMenuButtonX,nMenuButtonY, BGColor, ButtonClass)
	SettingsMenu = _
"<div id='divSettings' name='divSettings' style='color: " & BGColor & " ;background-color:" & BGColor & "; width: " & MenuW & "px; height: " & MenuH - 2 * TitleH & "px; position: absolute; right: -" & MenuH & "px; top: " & TitleH & "px;'>" &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: relative; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & BGColor & "; background-color: Transparent; width: " & MenuW & "px;"">" & _
			"<tbody>" & _
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & TitleH & """ width=""" & MenuW & """>" & _
					"</td>"&_
				"</tr>" &_
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""center"" class=""oa1"" height=""" & TitleH & """ width=""" & MenuW & """>" & _
						"<button class=""" & ButtonClass & """ style='width:" & nMenuButtonX & ";height:" & nMenuButtonY & ";' " &_
						"id='s_Button0' name='SET_TOPOLOGY' onclick=document.all('ButtonHandler').value='SET_TOPOLOGY';>CFG Template Name</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: none; background-color: Transparent;""align=""center"" class=""oa1"" height=""" & Int(TitleH/4) & """ width=""" & MenuW & """>" & _
						"<button class=""" & ButtonClass & """ style='width:" & nMenuButtonX & ";height:" & nMenuButtonY & ";' " &_
						"id='s_Button1'  name='SET_CONNECTIVITY' onclick=document.all('ButtonHandler').value='SET_CONNECTIVITY';>LAB Connectivity</button>" & _	
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
					"<td style=""border-style: None; background-color: Transparent;""align=""center"" class=""oa1"" height=""" & TitleH & """ width=""" & MenuW & """>" & _
    					"<button class=""" & ButtonClass & """ style='width:" & nMenuButtonX & ";height:" & nMenuButtonY & ";' " &_
						"id='s_Button2'  name='SET_FOLDER' AccessKey='E' onclick=document.all('ButtonHandler').value='SET_FOLDER';>Folder Settings</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""center"" class=""oa1"" height=""" & TitleH & """ width=""" & MenuW & """>" & _
    					"<button class=""" & ButtonClass & """ style='width:" & nMenuButtonX & ";height:" & nMenuButtonY & ";' " &_
						"id='s_Button3'  name='SET_CRT' AccessKey='E' onclick=document.all('ButtonHandler').value='SET_CRT';>SecureCRT</button>" & _
					"</td>"&_
				"</tr>" &_
				"<tr>" &_
    				"<td style=""border-style: none; background-color: Transparent;""align=""center"" class=""oa1"" height=""" & TitleH & """ width=""" & MenuW & """>" & _
						"<button class=""" & ButtonClass & """ style='width:" & nMenuButtonX & ";height:" & nMenuButtonY & ";' " &_
						"id='s_Button4'  name='SET_OTHER' onclick=document.all('ButtonHandler').value='SET_OTHER';>Other Settings</button>" & _	
					"</td>"&_
				"</tr>" &_
		    "</tbody></table>" &_
	 "</div>"
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
'----------------------------------------------
'Function  GetValueOld(nInd)
'----------------------------------------------
Function GetValueOld(nInd)
Dim i
	GetValueOld = UNKNOWN
	If IsNumeric(nInd) Then 
		GetValueOld = Trim(Split(vOld_Settings(nInd),"=")(1))
	Else 
		For i=0 to UBound(vParamNames) - 1
		    If vParamNames(i) = nInd Then 
			    On Error Resume Next
				GetValueOld = Trim(Split(vOld_Settings(i),"=")(1))
				if Err.Number > 0 Then MsgBox "nInd = " & nInd & chr(13) & "vOld_Settings(i)=" & vOld_Settings(i)
				On Error GoTo 0 
				Exit For
			End If 
		Next 
	End If
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
'------------------------------------------------
'    SETTINGS DIALOG FORM 
'------------------------------------------------
Function IE_PromptForSettings_Internal(nMenu, ByRef g_objIE, ByRef vIE_Scale, byRef vPlatforms, nDebug)
	Dim g_objShell, objShellApp, objFSO, objRegEx
	Dim nInd
	Dim nRatioX, nRatioY, nFontSize_10, nFontSize_12, nButtonX, nButtonY, nA, nB
    Dim intX
    Dim intY
	Dim nCount
	Dim strLogin
	Dim nLine, nService, nFlavor, nTask, nPlatform
	Dim vvMsg(8,3)
	Dim objFile, objCfgFile
	Dim objWMIService, IPConfigSet
	Const MAX_PARAM = 40
	Const MAX_BW_PROFILES = 30
	Dim objFolder, objForm, colFiles, strFile
	Call TrDebug ("IE_PromptForInputPullDown: OPEN MAIN CONFIG LOADER FORM ", "", objDebug, MAX_LEN, 3, nDebug)	
	Set objForm = CreateObject("Shell.Application")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Redim vOld_Settings(Ubound(vSettings))
	Dim BackGroundColor,ButtonColor ,TitleBgColor,TitleTextColor,ParentWindowH,ParentWindowW
	Dim WindowH, WindowW
	'
	' GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	IE_Window_Title = strTitle
	'
	' SCREEN RESOLUTION
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	IE_PromptForSettings_Internal = True	
	'
	'  READ LOCAL IPCONFIG
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	'----------------------------------------
	' MAIN VARIABLES OF THE GUI FORM
	'----------------------------------------
	If nRatioX > 1 Then nRatioX = 1 : nRatioY = 1 End If
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	nBottom = Round(10 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellH = Round(24 * nRatioY,0)
	LoginTitleW = Round(800 * nRatioX,0)
	nLeft = Round(20 * nRatioX,0)
	nTab = Round(40 * nRatioX,0)
	CellW = LoginTitleW
	LoginTitleH = Round(40 * nRatioY,0)
	nSaveW = nLeft + nButtonX
	nScoreW = 3 * nSaveW
	nColumn = Int(LoginTitleW/3)	
	nNameW = Int((LoginTitleH - nColumn)/3)
	'
	'	Calculate Settings window/div size
	vLine = Array(15,15,15,15)
	vTitle = Array("Test Bed Topology and Test Cases Parameters",_
	                 "FTP and Connectivity settings",_
                     "SecureCRT Settings",_
                     "MefCfgLoader Settings")
	nLine = vLine(nMenu)
	WindowH = 2 * LoginTitleH + cellH * nLine + nButtonY + nBottom
	WindowW = LoginTitleW
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)
    ParentWindowH = g_objIE.height
	ParentWindowW =	g_objIE.width
	vMenuSettings = Array(Array(12,7,8,9,13,14),Array(0,1,2,3,4),Array(6),Array(5,10,11,17))
    '
	' DRAW THE TITLE OF THE  FORM   		
	'-----------------------------------------------------------------
	nLine = 1
	    strHTMLbody = 	"<input name='ButtonHandler_my' type='hidden' value='Nothing Clicked Yet'>" &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: " & HttpBgColor5 & "; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""" &_
			"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW & """>" & _
				"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
				";font-weight: normal;font-style: italic;"">"&_
				"&nbsp;&nbsp;"& vTitle(nMenu) & "</span></p>"&_
			"</td>" &_
		"</tr></tbody></table>"
	Select Case nMenu
	    Case 0
			'-----------------------------------------------------------------
			' TOPOLOGY AND TEST CASE PARAMS
			'-----------------------------------------------------------------
			cTitle = "Topology and Test Series"
			strHTMLbody = strHTMLbody &_
				"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
				"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
				"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
					"<tbody>" &_
						"<tr>" & _
							"<td colspan = ""3"" style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
							";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
								";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
							"</td>" & _
						"</tr>"
			nSetting = 12
			nLine = nLine + 1
			strHTMLbody = strHTMLbody &_
				"<tr>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
					"</td>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
						"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
						" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
						"type=text > " &_
					"</td>" &_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<button class = ""low_button"" style='width:" & 1.5 * nButtonX &_
						";height:" & Int(nButtonY/2) &_
						"px' name='Edit_PARAM' onclick=document.all('ButtonHandler_my').value='EDIT_PARAM';>Edit File</button>" & _	
					"</td>" &_
				"</tr>" &_
				"<tr>"&_
					"<td colspan =""2"" style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & 3 * Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
					"</td>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<button class = ""low_button"" style='width:" & 1.5 * nButtonX &_
						";height:" & Int(nButtonY/2) &_
						"px' name='Select_PARAM' onclick=document.all('ButtonHandler_my').value='SELECT_PARAM';>Import File</button>" & _	
					"</td>" &_
				"</tr>" &_
				"<tr>"&_
					"<td colspan =""2"" style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & 3 * Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
					"</td>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<button class = ""low_button"" style='width:" & 1.5 * nButtonX &_
						";height:" & Int(nButtonY/2) &_
						"px' name='Select_PARAM' onclick=document.all('ButtonHandler_my').value='RELOAD_TOPOLOGY';>Reload</button>" & _	
					"</td>" &_
				"</tr>" &_
			"</tbody></table>"			
			nLine = nLine + 4
			'
			' FOLDERS TITLE
			cTitle = "Folder Settings"
			strHTMLbody = strHTMLbody &_
				"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
				"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
				"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
					"<tbody>" &_
						"<tr>" & _
							"<td colspan = ""3"" style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
							";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
								";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
							"</td>" & _
						"</tr>"
			'
			' FOLDERS PARAMETERS:
			For nSetting = 7 to 9
			nLine = nLine + 1
			strType = "text"
			strHTMLbody = strHTMLbody &_
				"<tr>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
					"</td>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
						"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
						" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						"; background-color: " & HttpBgColor4  & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
						"type=" & strType & " > " &_
					"</td>" &_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<button class = ""low_button"" style='width:" & 1.5 * nButtonX & "; height:" & Int(nButtonY/2) & "px;' " &_
						"name='Edit_Folder'" & nSetting & " onclick=document.all('ButtonHandler_my').value='Folder_" & nSetting & "' >Select Folder</button>" & _	
					"</td>" &_
				"</tr>"
			Next
			strHTMLbody = strHTMLbody & "</tbody></table>"
			nLine = nLine + 2
			'
			' PLATFORM TITLE
			cTitle = "Platform under test"
			strHTMLbody = strHTMLbody &_
				"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
				"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
				"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
					"<tbody>" &_
						"<tr>" & _
							"<td colspan=""3"" style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
							";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
								";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
						"</tr>"
			nLine = nLine + 1
			'
			' PLATFORM PARAMETERS:
			strType = "text"
			strHTMLbody = strHTMLbody &_
				"<tr>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(13) & "</p>" &_
					"</td>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
					"</td>" &_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<select name='Platform_Name' id='Platform_Name' class = ""regular""" &_
							"style=""width: "& 1.5 * nButtonX & "; border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
							" onchange=document.all('ButtonHandler_my').value='SelectPlatform';" &_
							"type=text > " &_
						"</select>" &_
					"</td>" &_
				"</tr>"
			nSetting = 14
			strHTMLbody = strHTMLbody &_
				"<tr>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(14) & "</p>" &_
					"</td>"&_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_
						"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
						";font-weight: normal;font-style: italic;"">"&_
						"&nbsp;&nbsp;Config. name: [Service Type]-[TC#]-<span style=""font-weight: bold;"">[Prefix]</span>-[L|R].conf</span></p>"&_
					"</td>" &_
					"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
						"<input name=Settings_Param_" & nSetting & " value='' style=""width: " & 1.5 * nButtonX & "text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
						" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
						"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=15 " &_
						"type=" & strType & " > " &_
					"</td>" &_
				"</tr>"
			strHTMLbody = strHTMLbody &_
					"</tbody></table>"
			nLine = nLine + 2
			
		Case 1	
		    '
			' CONNECTIVITY SETTINGS TITLE
			cTitle = "Connectivity Settings"
			strHTMLbody = strHTMLbody &_
				"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
				"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
				"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
					"<tbody>" &_
						"<tr>" & _
							"<td colspan = ""3"" style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
							";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
								";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
							"</td>" & _
						"</tr>"
			For nSetting = 0 to 4
				nLine = nLine + 1
				strType = "text"
				If InStr(vParamNames(nSetting),"assword") Then strType = "password" End If
				strHTMLbody = strHTMLbody &_
					"<tr>"&_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
						"</td>"
				Select Case nSetting
					Case 2
							strHTMLbody = strHTMLbody &_
								"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
									"<select name='Adapter_Name' id='Adapter_Name' class = ""regular""" &_
										"style=""width: "& Int(LoginTitleW/2) - nTab & ";border: none ; outline: none; text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
										"font-family: 'Helvetica'; color: " & HttpTextColor2 &_
										"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" size='1'" & _
										" onchange=document.all('ButtonHandler_my').value='SelectAdapter';" &_
										"type=text > " &_
							        "</select>" &_
								"</td>"
					Case Else
							strHTMLbody = strHTMLbody &_
									"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center""></td>"
				End Select
				strHTMLbody = strHTMLbody &_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
							"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: right; font-size: " & nFontSize_10 & ".0pt;" &_ 
							" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=18 maxlength=18 " &_
							"type=" & strType & " > " &_
						"</td>" &_
					"</tr>" 
			    Next
				strHTMLbody = strHTMLbody &	"</tbody></table>"
				nLine = nLine + 2
		Case 2
					'-----------------------------------------------------------------
					' FOLDERS TITLE
					'-----------------------------------------------------------------
					cTitle = "Folder Settings"
					strHTMLbody = strHTMLbody &_
						"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
						"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
						"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
							"<tbody>" &_
								"<tr>" & _
									"<td colspan = ""3"" style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
									";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
										"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
										";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
									"</td>" & _
								"</tr>" &_
								"<tr><td colspan = ""3"" style="" border-style: None;"" height=""" & cellH & """></td></tr>"
				'-----------------------------------------------------------------
				' FOLDERS PARAMETERS:
				'-----------------------------------------------------------------
			'	nColumn = Int(nScoreW/3)
				nSetting = 6
					nLine = nLine + 1
					strType = "text"
					BgTextColor = HttpBgColor4 : ButtonDisabled = ""
					If nSetting = 5 Then ButtonDisabled = "disabled" : BgTextColor = HttpBgColor1  End If			
					strHTMLbody = strHTMLbody &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & BgTextColor & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
								"type=" & strType & " > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """ align=""center"">" &_									
								"<button class = ""low_button"" style='width:" & 2 * nButtonX &_
								";height:" & Int(nButtonY/2) &_
								"px' name='Edit_Folder'" & nSetting & " onclick=document.all('ButtonHandler_my').value='Folder_" & nSetting & "'; " & ButtonDisabled & ">Edit Folder</button>" & _	
							"</td>" &_
						"</tr>"
				strHTMLbody = strHTMLbody &_
						"</tbody></table>"
				nLine = nLine + 3
		Case 3
					'-----------------------------------------------------------------
					' SECURECRT SESSIONS TITLE
					'-----------------------------------------------------------------
					cTitle = "SecureCRT Sessions"
					strTitleCell = 	"<td style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
									";"" align=""right"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """>" & _
										"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
										";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;&nbsp;&nbsp;"

					strHTMLbody = strHTMLbody &_
						"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
						"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
						"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
							"<tbody>" &_
								"<tr>" & _
									"<td style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
									";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
										"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
										";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
									"</td>" &_
									strTitleCell & "Session Folder" & "</p></td>" &_
									strTitleCell & "Session Name" & "</p></td>" &_
									strTitleCell & "Host Name" & "</p></td>" &_
									strTitleCell & "Session Login" & "</p></td>" &_
								"</tr>" &_
								"<tr><td colspan = ""5"" style="" border-style: None;"" height=""" & cellH & """></td></tr>" 								
				'-----------------------------------------------------------------
				' SECURECRT SESSIONS PARAMETERS:
				'-----------------------------------------------------------------
				For nSetting = 10 to 11
					nLine = nLine + 1
					BgTextColor = HttpBgColor4
					If nSetting = 11 Then BgTextColor = HttpBgColor1  End If			
					strHTMLbody = strHTMLbody &_
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & BgTextColor & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & HttpBgColor4 & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
								"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
								" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								"; background-color: " & BgTextColor & "; font-weight: Normal;"" AccessKey=i size=15 maxlength=128 " &_
								"type=text > " &_
							"</td>" &_
						"</tr>"
				Next
				strHTMLbody = strHTMLbody &_				
						"<tr>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
								"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
								";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & HIDE_CRT & "</p>" &_
							"</td>"&_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
									"<input type=checkbox name='Hide_CRT' style=""color: " & HttpTextColor2 & ";""" & _
									" onclick=document.all('ButtonHandler_my').value='HIDE_CRT';" &_
									"value='Display'>" &_
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
							"</td>" &_
							"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(3 * LoginTitleW/16) & """ align=""center"">" &_									
							"</td>" &_
						"</tr>"

				strHTMLbody = strHTMLbody &_
						"</tbody></table>"
				nLine = nLine + 4
				'
				' FOLDERS TITLE
				cTitle = "Folder Settings"
				strHTMLbody = strHTMLbody &_
					"<table border=""1"" cellpadding=""1"" cellspacing=""1"" width=""" & LoginTitleW & """ valign=""middle""" &_ 
					"style="" position: absolute; top: " & LoginTitleH + nLine * cellH & "px; left: 0px;" &_
					"border-collapse: collapse; border-style: none ; background-color: " & HttpBgColor6 & "'; width: " & LoginTitleW & "px;"">" & _
						"<tbody>" &_
							"<tr>" & _
								"<td colspan = ""3"" style="" border-style: solid;background-color: " & HttpBgColor6 & "; border-color: " & HttpBgColor6 &_
								";"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
									"<p style=""text-align: Left; font-size: " & nFontSize_12 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor1 &_
									";font-weight: bold;font-style: bold;"">&nbsp;&nbsp;" & cTitle & "</p>" &_
								"</td>" & _
							"</tr>" & _
							"<tr><td colspan = ""3"" style="" border-style: None;"" height=""" & cellH & """></td></tr>"
				'
				' FOLDERS PARAMETERS:
				nSetting = 5
				nLine = nLine + 1
				strType = "text"
				strHTMLbody = strHTMLbody &_
					"<tr>"&_
						"<td style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<p style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							";font-weight: normal;font-style: normal;"">&nbsp;&nbsp;" & vParamNames(nSetting) & "</p>" &_
						"</td>"&_
						"<td colspan = ""2"" style="" border-style: None;"" class=""oa2"" height=""" & cellH & """ width=""" & Int(LoginTitleW/2) & """ align=""center"">" &_									
							"<input name=Settings_Param_" & nSetting & " value='' style=""text-align: Left; font-size: " & nFontSize_10 & ".0pt;" &_ 
							" border-style: none; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
							"; background-color: " & BgTextColor & "; font-weight: Normal;"" AccessKey=i size=50 maxlength=128 " &_
							"type=" & strType & " > " &_
						"</td>" &_
					"</tr>" &_
				"</tbody></table>"
				nLine = nLine + 3
	End Select 
	'
	'   EXIT BUTTON
	strHTMLbody = strHTMLbody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: " & HttpBgColor2 &_
		"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button  class=""ok_button""" &_
						"style='width:" & 2 * nButtonX & ";height:" & nButtonY & ";' " &_
						"name='SAVE' onclick=document.all('ButtonHandler_my').value='SAVE';><u>S</u>ave Settings</button>" & _	
					"</td>"&_
					"<td colspan=""2"" style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
					"</td>"&_
					"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button class=""ok_button"" style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor3 & "; width:" &_
						2 * nButtonX & ";height:" & nButtonY & ";font-size: " & nFontSize_12 & ".0pt;" &_
						"px;' name='EXIT' onclick=document.all('ButtonHandler_my').value='Cancel';><u>E</u>xit</button>" & _	
					"</td>"&_
				"</tr></tbody></table>"

	'-----------------------------------------------------------------
	' HTML Form Parameaters
	'-----------------------------------------------------------------
	'
	'  Load Message board form
	g_objIE.Document.getElementById("divMessage").innerHTML  = strHTMLBody
	'
	'  Message Board Properties
	g_objIE.Document.getElementById("divMessage").style.zindex  = "999"
	g_objIE.Document.getElementById("divShadow").style.zindex  = "998"
	g_objIE.Document.getElementById("divMessage").style.position  = "absolute"
	g_objIE.Document.getElementById("divMessage").style.top = (ParentWindowH - WindowH)/2 & "px"
	g_objIE.Document.getElementById("divMessage").style.left = (ParentWindowW - WindowW)/2 & "px"
	g_objIE.Document.getElementById("divMessage").style.backgroundColor = "grey"
	g_objIE.Document.getElementById("divMessage").style.color  = HttpTextColor2	
	g_objIE.Document.getElementById("divMessage").style.fontsize  = "20px"
	g_objIE.Document.getElementById("divMessage").style.fontfamily = "arial narrow"	
	g_objIE.Document.getElementById("divMessage").style.height  = WindowH & "px"	
	g_objIE.Document.getElementById("divMessage").style.width  = WindowW & "px"
	g_objIE.Document.getElementById("divMessage").style.boxShadow="10px 10px 10px #101010"
	
	'-----------------------------------------------------------------
	'	SET DEFAULT PARAMETERS
	'-----------------------------------------------------------------
	Select Case nMenu
		Case 0 ' Topology and platforms
		    ' Load values for platform names
			For nInd = 0 to Ubound(vPlatforms) - 1
				g_objIE.document.getElementById("Platform_Name").Length = nInd + 1
				g_objIE.document.getElementById("Platform_Name").Options(nInd).text = Split(vPlatforms(nInd),",")(0)
			Next
			If GetValue(PLATFORM_NAME) <> "Unknown" Then nIndex = GetObjectLineNumber( vPlatforms, UBound(vPlatforms), GetValue(PLATFORM_NAME)) - 1 Else nIndex = 0 End If
			g_objIE.document.getElementById("Platform_Name").selectedIndex = nIndex
			g_objIE.Document.All("Settings_Param_" & 14).Value = GetValue(PLATFORM_INDEX)
			g_objIE.Document.All("Settings_Param_" & 7).Value = GetValue(CONFIGS_FOLDER)
			g_objIE.Document.All("Settings_Param_" & 8).Value = GetValue(Orig_Folder)
		    g_objIE.Document.All("Settings_Param_" & 12).Value = GetValue(CONFIGS_PARAM)
		Case 1 ' Connectivity
			nInd = 0
			'  Load network adapters list
			For Each IPConfig in IPConfigSet
				g_objIE.document.getElementById("Adapter_Name").Length = nInd + 1
				g_objIE.document.getElementById("Adapter_Name").Options(nInd).text = IPConfig.Description
				g_objIE.document.getElementById("Adapter_Name").Options(nInd).value = IPConfig.IPAddress(0)
				if IPConfig.IPAddress(0) = GetValue(FTP_IP) _ 
					Then g_objIE.document.getElementById("Adapter_Name").selectedIndex = nInd
				nInd = nInd + 1
			Next
			' Add loopback to the list
			strLine = strLine &	"<option value=127.0.0.1>Loopback</option>"
			g_objIE.document.getElementById("Adapter_Name").Length = nInd + 1
			g_objIE.document.getElementById("Adapter_Name").Options(nInd).text = "Loopback"
			g_objIE.document.getElementById("Adapter_Name").Options(nInd).Value = "127.0.0.1"
			if g_objIE.document.getElementById("Adapter_Name").Options(nInd).Value = GetValue(FTP_IP) _ 
				Then g_objIE.document.getElementById("Adapter_Name").selectedIndex = nInd
			' Load other connectivity params
			For nInd = 0 to 4
				g_objIE.Document.All("Settings_Param_" & nInd).Value = GetValue(nInd)
			Next		
		Case 2 ' Folders  
			g_objIE.Document.All("Settings_Param_" & 6).Value = GetValue(WORK_FOLDER)
		Case 3 ' SecureSRT
			g_objIE.Document.All("Settings_Param_" & 5).Value = GetValue(SECURECRT_FOLDER)
		    If GetValue(HIDE_CRT) = "True" Then g_objIE.Document.All("Hide_CRT").Checked = True
			For i = 0 to UBound(Split(Split(vSettings(10),"=")(1),","))
				Select Case i
				   Case 0, 3
					  g_objIE.Document.All("Settings_Param_11")(i).Value = "Same as above"
					  g_objIE.Document.All("Settings_Param_10")(i).Value = Split(GetValue(10),",")(i)
				   Case 1, 2
					  g_objIE.Document.All("Settings_Param_10")(i).Value = Split(GetValue(10),",")(i)
					  g_objIE.Document.All("Settings_Param_11")(i).Value = Split(GetValue(11),",")(i)
				End Select 
			Next
	End Select
	'
	'  Main Cycle of the Message Form 
	Do
        On Error Resume Next
		    Err.clear
            szNothing = g_objIE.Document.All("ButtonHandler_my").Value
            if Err.Number <> 0 then exit function
        On Error Goto 0    
        ' Check to see which buttons have been clicked, and address each one
        ' as it's clicked.
        Select Case szNothing
		    Case "SelectPlatform"
			            nPlatform = g_objIE.document.getElementById("Platform_Name").selectedIndex
			            g_objIE.Document.All("Platform_Index").Value = Split(vPlatforms(nPlatform),",")(1)
						g_objIE.Document.All("ButtonHandler_my").Value = "Nothing Selected"
			Case "SelectAdapter"
			            ' nAdapter = g_objIE.document.getElementById("Adapter_Name").selectedIndex
			            g_objIE.Document.All("Settings_Param_2").Value = g_objIE.document.getElementById("Adapter_Name").Value
						g_objIE.Document.All("ButtonHandler_my").Value = "Nothing Selected"   
			Case "Cancel"
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings_Internal = 0
						'Call FocusToParentWindow(strPID)
						exit function
			Case "Update"
			            Do 
							If InStr(strUpdate,"PARAM") Then 
								If ValidateParamFile(g_objIE.Document.All("Settings_Param_12").Value, False) Then 
									strFileParam = g_objIE.Document.All("Settings_Param_12").Value
									g_objIE.Document.All("ButtonHandler_my").Value = "Exit_and_Reload"
									Exit Do
								Else 
									Call CopyArray(vOld_Settings, vSettings, nDebug)
									MsgBox "Topology file has a wrong format." & chr(13) & "Select another topology file..."
									g_objIE.Document.All("ButtonHandler_my").Value = "None"
									Exit Do								
								End If 
							End If
							If InStr(strUpdate,"ZEROCFG") Then 
								'   CREATE ZERO CONFIG FILES
								Old_Platform_Index = GetValue(PLATFORM_INDEX)
								Call Create_Zero_Configs(vSettings, Old_Platform_Index)
							End If
							'   CREATE FTP USER AND HOME FOLDERS IN FELIZILLA SETTINGS
							If InStr(strUpdate,"FTP") Then 
								If FileZilla_Installed Then 
									Set objShellApp = CreateObject("Shell.Application")
									objShellApp.ShellExecute "wscript", """" & VBScript_FTP_User & """" &_
																		" """ & strDirectoryConfig & """ " &_
																		GetValue(FTP_User) & " " &_
																		Split(vOld_Settings(3),"=")(1) & " " &_
																		GetValue(FTP_Password), "", "runas", 1
									Set objShellApp = Nothing
								End If														   
							End If  
							'   CREATE SECURE CRT SESSIONS						
							If InStr(strUpdate,"CRT") Then 
								If SecureCRT_Installed Then Call Create_CRT_Sessions(vSettings, vOld_Settings)						   
							End If 
							g_objIE.Document.All("ButtonHandler_my").Value = "Save_and_Exit"   
						Exit Do
						Loop
			Case "Save_and_Exit"
						For Each nSettings in vMenuSettings(nMenu)
							nInd = CInt(nSettings)
							If GetValue(nInd) <> GetValueOld(nInd) Then 
								Call FindAndReplaceStrInFile(strFileSettings, vParamNames(nInd), vSettings(nInd), 0)
								Call FindAndReplaceStrInFile(strFileParam, vParamNames(nInd), vSettings(nInd), 0)
						    End If 
							'vSettings(nInd) = NormalizeStr(vSettings(nInd),vDelim)
						Next 
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings_Internal = 1
						'Call FocusToParentWindow(strPID)
						exit function
			Case "Exit_and_Reload"
						If GetValue(CONFIGS_PARAM) <> GetValueOld(CONFIGS_PARAM) Then 
							Call FindAndReplaceStrInFile(strFileSettings, vParamNames(12), vSettings(12), 0)
						End If
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings_Internal = 2
						'Call FocusToParentWindow(strPID)
						exit function			           
			Case "Exit_And_Close_Wscript"
						Set objForm = Nothing
						Set objFolder = Nothing
						Set g_objShell = Nothing
						Set objFSO = Nothing
						IE_PromptForSettings_Internal = -1
						exit function
			Case "Folder_5"
						strFolder = g_objIE.Document.All("Settings_Param_5").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_5").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
			Case "Folder_6"
						strFolder = g_objIE.Document.All("Settings_Param_6").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_6").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
			Case "Folder_7"
						strFolder = g_objIE.Document.All("Settings_Param_7").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_7").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
			Case "Folder_8"
						strFolder = g_objIE.Document.All("Settings_Param_8").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_8").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
			Case "Folder_9"
						strFolder = g_objIE.Document.All("Settings_Param_9").Value
						Set objFolder = objForm.BrowseForFolder(0, "Choose Local Folder", 0, strFolder)
						If Not objFolder Is Nothing Then
							strFolder = objFolder.self.path
							g_objIE.Document.All("Settings_Param_9").Value = strFolder
						End If
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
												
		    Case "SAVE" 
						g_objIE.Document.All("ButtonHandler_my").Value = "Update"
						strUpdate = "Update"
						Call CopyArray(vSettings, vOld_Settings, nDebug)
						Set objRegEx = CreateObject("VBScript.RegExp")
						objRegEx.Global = False
						Do
							Select Case nMenu
							    Case 0
										'
										'  UPDATE vSETTINGS 
										For Each nSettings in vMenuSettings(nMenu)
											nInd = CInt(nSettings)
											'  Do simple folder check
											objRegEx.Pattern =".+Folder[$\s]"
											If objRegEx.Test(vParamNames(nInd)) Then 
												If Not objFSO.FolderExists(g_objIE.Document.All("Settings_Param_" & nInd).Value) Then 
													MsgBox "Can't find folder: " & chr(13) & g_objIE.Document.All("Settings_Param_" & nInd).Value & chr(13) & "Try again..."
													Call CopyArray(vOld_Settings, vSettings, nDebug)
													Exit Do
												End If
											End If 
											'  Do simple file check
											objRegEx.Pattern =".+File[$\s]"
											If objRegEx.Test(vParamNames(nInd)) Then 
												If Not objFSO.FileExists(g_objIE.Document.All("Settings_Param_" & nInd).Value) Then 
													MsgBox "Can't find file: " & chr(13) & g_objIE.Document.All("Settings_Param_" & nInd).Value & chr(13) & "Try again..."
													Call CopyArray(vOld_Settings, vSettings, nDebug)
													Exit Do
												End If
											End If 
						                    '  copy new settings value from form
											Select Case nInd
												Case 13
													nPlatform = g_objIE.document.getElementById("Platform_Name").selectedIndex
													Call SetValue(nInd,g_objIE.document.getElementById("Platform_Name").Options(nPlatform).text, False)
												Case Else
													Call SetValue(nInd,g_objIE.Document.All("Settings_Param_" & nInd).Value, False)
											End Select
										Next
										If GetValue(7) <> GetValueOld(7) Then 
										   strUpdate =  strUpdate & "_FTP" 
										   Exit Do
										End If 
										If GetValue(14) <> GetValueOld(14) Then 
										   strUpdate =  strUpdate & "_ZEROCFG" 
										   Exit Do
										End If 
										If GetValue(CONFIGS_PARAM) <> GetValueOld(CONFIGS_PARAM) Then 
										   strUpdate =  strUpdate & "_PARAM" 
										   Exit Do
                                        End If										
								Case 1
										'
										'  UPDATE vSETTINGS AND WRITE SETTINGS TO FILE
										For Each nSettings in vMenuSettings(nMenu)
											nInd = CInt(nSettings)
											' Do simple check of the IP address
											If InStr(vParamNames(nInd),"Node IP") Then 
												objRegEx.Pattern ="\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}"
												If Not objRegEx.Test(g_objIE.Document.All("Settings_Param_" & nInd).Value)	Then
													MsgBox "Wrong IP address format: " & g_objIE.Document.All("Settings_Param_" & nInd).Value &_
															chr(13) & "Use the following format for management IP address: "  &  "A.B.C.D/Prefix"
													Call CopyArray(vOld_Settings, vSettings, nDebug)															
													g_objIE.Document.All("ButtonHandler_my").Value = "None"
													Exit Do
												End If						
											End If 
											Call SetValue(nInd,g_objIE.Document.All("Settings_Param_" & nInd).Value, False)
										Next
											
										If GetValue(3) <> GetValueOld(3) or _
										   GetValue(4) <> GetValueOld(4) Then 
										   strUpdate =  strUpdate & "_FTP"
										   Exit Do
										End If 

										If GetValue(0) <> GetValueOld(0) or _
										   GetValue(1) <> GetValueOld(1) Then 
										   strUpdate =  strUpdate & "_CRT"
										   Exit Do
										End If 
								Case 2
										'
										'  UPDATE vSETTINGS AND WRITE SETTINGS TO FILE
										Call SetValue(6,g_objIE.Document.All("Settings_Param_" & 6).Value, False)
										g_objIE.Document.All("ButtonHandler_my").Value = "Cancel" 
								Case 3
										For Each nSettings in vMenuSettings(nMenu)
											nInd = CInt(nSettings)
											'  Do simple folder check
											objRegEx.Pattern =".+Folder[$\s]"
											If objRegEx.Test(vParamNames(nInd)) Then 
												If Not objFSO.FolderExists(g_objIE.Document.All("Settings_Param_" & nInd).Value) Then 
													MsgBox "Can't find folder: " & chr(13) & g_objIE.Document.All("Settings_Param_" & nInd).Value & chr(13) & "Try again..."
													Call CopyArray(vOld_Settings, vSettings, nDebug)
													Exit Do
												End If
											End If 
											strLine = ""
											Select Case nInd
												Case 17
													If g_objIE.Document.All("Hide_CRT").Checked Then
														Call SetValue(nInd,"True", False)
														nWindowState = HideTerminal
													Else 
														nWindowState = ShowTerminal
														Call SetValue(nInd,"False", False)
													End If												
												Case 10, 11
													strLine = Trim(g_objIE.Document.All("Settings_Param_" & 10)(0).Value) & ","
													strLine = strLine & Trim(g_objIE.Document.All("Settings_Param_" & nInd)(1).Value) & ","
													strLine = strLine & Trim(g_objIE.Document.All("Settings_Param_" & nInd)(2).Value) & ","
													strLine = strLine & Trim(g_objIE.Document.All("Settings_Param_" & 10)(3).Value)
													Call SetValue(nInd,strLine, False)
												Case Else
													Call SetValue(nInd,g_objIE.Document.All("Settings_Param_" & nInd).Value, False)													
											End Select
										Next
										If GetValue(10) <> GetValueOld(10) or _
										   GetValue(11) <> GetValueOld(11) Then 
										   strUpdate =  strUpdate & "_CRT"
										   Exit Do
										End If 
							End Select
							Exit Do
					    Loop
			Case "RELOAD_TOPOLOGY"
						Call CopyArray(vSettings, vOld_Settings, nDebug)
						g_objIE.Document.All("ButtonHandler_my").Value = "Exit_and_Reload"
			Case "SELECT_PARAM"
						g_objIE.Document.All("Settings_Param_12").Value = BrowseForFile()
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
			Case "EDIT_PARAM"
						objEnvar.Run strEditor & " " & strFileParam
						' MsgBox "notepad.exe " & strFileParam
						g_objIE.Document.All("ButtonHandler_my").Value = "None"
			Case "SET_NODE"
                        g_objIE.Document.All("ButtonHandler_my").Value = "None"
        End Select
        
		WScript.Sleep 300
    Loop
End Function
'------------------------------------------------------------------------
'	Function CopyArray(ByRef Array1, ByRef Array2, nDebug)
'------------------------------------------------------------------------
Function CopyArray(ByRef Array1, ByRef Array2, nDebug)
Dim vDim(3), nDim
	nDim = 0
	vDim(0) = 0 : vDim(1) = 0 : vDim(2) = 0 
	On Error resume next
    Err.Clear
	Do While nDim < 3
	    vDim(nDim) = UBound(Array1,nDim + 1)
	    If Err.Number <> 0 Then Exit Do
		nDim = nDim + 1
	Loop
	Call TrDebug("nDim = " & nDim & " Error: " & Err.Description, "", objDebug, MAX_LEN , 1, nDebug)
	' nDim equals to number Array1 dimentions: 1, 2 or 3
	On Error goto 0
	CopyArray = nDim
    Select Case nDim
        Case 1
		    Redim Array2(UBound(Array1,1))
			For i = 0 to UBound(Array1,1) - 1 
			    Array2(i) = Array1(i)
				Call TrDebug("Array(" & i & ") = " & Array2(i), "", objDebug, MAX_LEN , 1, nDebug)
			Next
        Case 2
		    Redim Array2(UBound(Array1,1),UBound(Array1,2))
			For i = 0 to UBound(Array1,1) - 1 
			    For n = 0 to UBound(Array1,2) - 1 
					Array2(i,n) = Array1(i,n)
					Call TrDebug("Array(" & i & "," & n & ") = " & Array2(i,n), "", objDebug, MAX_LEN , 1, nDebug)
				Next
			Next
			
        Case 3		
	        Redim Array2(UBound(Array1,1),UBound(Array1,2),UBound(Array1,3))
			For i = 0 to UBound(Array1,1) - 1 
			    For n = 0 to UBound(Array1,2) - 1 
					For k = 0 to UBound(Array1,3) - 1 
						Array2(i,n,k) = Array1(i,n,k)
						Call TrDebug("Array(" & i & "," & n & "," & k & ") = " & Array2(i,n,k), "", objDebug, MAX_LEN , 1, nDebug)
					Next
				Next
			Next
	End Select 
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
'--------------------------------------------------
'   Function ValidateParamFile(strFile, bCheckVersion)
'--------------------------------------------------
Function ValidateParamFile(strFile, bCheckVersion)
Dim nLine, vFileLine
    ValidateParamFile = False
    If GetFileLineCountByGroup(strFile, vFileLine,"Description","","",0) = 0 Then Exit Function
    If GetObjectLineNumber(vFileLine, UBound(vFileLine), "MEFcfgLoader topology") = 0 Then Exit Function
    If bCheckVersion _ 
		Then If GetObjectLineNumber(vFileLine, UBound(vFileLine), "version 2") = 0 Then Exit Function
    ValidateParamFile = True
End Function
'-----------------------------------------------
' Function LoadTopology()
'-----------------------------------------------
Function LoadTopology(strFileParam,vNodes)
Dim nCount
	nCount = GetFileLineCountByGroup(strFileParam, vLines,"Node-Left","","",0)
	For nInd = 0 to nCount - 1
		Select Case Split(vLines(nInd),"=")(0)
				Case UNI_A
							vNodes(0,0) = Split(vLines(nInd),"=")(1)
				Case UNI_B
							vNodes(0,1) = Split(vLines(nInd),"=")(1)
				Case NNI
							vNodes(0,2) = Split(vLines(nInd),"=")(1)
		End Select
	next
	nCount = GetFileLineCountByGroup(strFileParam, vLines,"Node-Right","","",0)
	For nInd = 0 to nCount - 1
		Select Case Split(vLines(nInd),"=")(0)
				Case UNI_C
							vNodes(1,0) = Split(vLines(nInd),"=")(1)
				Case UNI_D
							vNodes(1,1) = Split(vLines(nInd),"=")(1)
				Case UNI_CC
							vNodes(1,2) = Split(vLines(nInd),"=")(1)
				Case UNI_DD
							vNodes(1,3) = Split(vLines(nInd),"=")(1)
				Case NNI
							vNodes(1,4) = Split(vLines(nInd),"=")(1)
		End Select
	Next
	'
	'  Read Default UNI/EVC parameters
	nCount = GetFileLineCountByGroup(strFileParam, vLines,"Default_Values","","",0)
	If nCount = 0 Then 
		CBSo = 1500
		CIRo = 1500		
		CBSdef = 13698
		PBSdef = 2 * CBSdef
		PBSo = CBSdef
		MTU = 9192
	Else 
		For nInd = 0 to nCount - 1
			Select Case Split(vLines(nInd),"=")(0)
					Case "CIR0"
								CIRo = Split(vLines(nInd),"=")(1)
								PIRo = CIRo
					Case "CBS0"
								CBSo = Split(vLines(nInd),"=")(1)
					Case "MTU"
								MTU = Split(vLines(nInd),"=")(1)
					Case "CBS"
								CBSdef = Split(vLines(nInd),"=")(1)
								PBSdef = 2 * CBSdef
								PBSo = CBSdef
			End Select
		Next
    End If		
End Function


'------------------------------------------------------------------------------
' Function GetTestCases
'------------------------------------------------------------------------------
Function GetTestCases(ByRef vSvc, ByRef vFlavors, nDebug)
Dim i, n, TaskID, strTaskLine, Platform 
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = False
    objRegEx.Pattern = "^\d-?\d?"
    Platform = GetValue(PLATFORM_INDEX)
	For i = 0 to UBound (vSvc,1) - 1
		Call TrDebug("Folder: " & strDirectoryConfig & "\" & vSvc(i,1) , "", objDebug, MAX_LEN, 3, nDebug)
	    For n = 0 to vSvc(i,0) - 1
		    strTaskLine = ""
			Set objFolder = objFSO.GetFolder(strDirectoryConfig & "\" & vSvc(i,1) )
			For Each objFile in objFolder.Files
				FileName = Split(objFile.Name,".")(0)
				If InStr(FileName,vFlavors(i,n,0))<>0 and InStr(FileName,Platform & "-l")<>0 Then 
					Call TrDebug("FileName: " & FileName , "", objDebug, MAX_LEN, 1, nDebug)
				    TaskID = Split(FileName,vFlavors(i,n,0) & "-")(1)
					TaskID = Split(TaskID,"-" & Platform)(0)
					If objRegEx.Test(TaskID) Then strTaskLine = strTaskLine & TaskID & ","
				End If 
			Next
			If Len(strTaskLine) > 0 Then strTaskLine = Left(strTaskLine,Len(strTaskLine)-1)
			vFlavors(i,n,1) =  strTaskLine
			Call FindAndReplaceStrInFile(strFileParam, vFlavors(i,n,0), vFlavors(i,n,0) & ":" & vFlavors(i,n,1), 0)			
			Set objFolder = Nothing
			Call TrDebug("Test Series: " & strTaskLine , "", objDebug, MAX_LEN, 1, nDebug)			
		Next
	Next
End Function



