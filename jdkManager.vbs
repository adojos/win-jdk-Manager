'###################################################################################################
'# SCRIPT NAME: jdkManager.vbs
'#
'# DESCRIPTION:
'# VBScript (WMI,WScript) Utility with Command Line Interface. Enables to view all 
'# installed JDK/JRE versions [32bit/64bit]. Easily view and re-point Env Variables.
'#
'# NOTES:
'# 32bit Location - HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
'# 64bit Location - HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\
'# Registry Enumeration uses 'WMI' & Environment Variables functions use 'WScript.Shell'
'# Since the WSH Shell has no Enumeration functionality, you cannot use the WSH Shell 
'# object to parse Registry "tree" unless you know the exact name of every subkey.
'# The WMI Classes used are 'StdRegProv' and 'Wim32_Environment'
'# Do not use WMI Class 'Win32_Product' as this class has known issues documented (KBArticle 794524)
'# https://support.microsoft.com/en-us/help/974524/event-log-message-indicates-that-the-windows-installer-reconfigured-al
'#
'# PLATFORM: Win7/8/Server | PRE-REQ: Script/Admin Privilege
'# LAST UPDATED: Wed, 17 Sept 2019 | AUTHOR: Tushar Sharma
'##################################################################################################




If WScript.Arguments.length = 0 Then
   Set objShell = CreateObject("Shell.Application")
   objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 3
      WScript.Quit
End If   


'###########################################################################

Dim dictInstalledJDKKRE
Dim dictHomeVars
Dim dictPathVars
Dim strSelectedOption

Const installedjava = "installedjava"
Const homevars = "homevars"
Const pathvars = "pathvars"

Const JDK = "JDK"
Const JRE = "JRE"
Const strJavaHome = "JAVA_HOME"
Const strJreHome = "JRE_HOME"
Const strPathEnvVar = "Path"
Const strJavaHomePathEnvVar = "JavaHomePathEnvVar"
Const NoPathVars = "NoPathVars"
Const NoJavaPathVars = "NoJavaPathVars"
Const strAllEnvVar = "strAllEnvVar"
Const strInvalid = "invalid"
arrPathTypes = Array("javapath", "%JAVA_HOME%\bin", "jdk", "jre")

Const FoundJdkJre = "FoundJdkJre"
Const FoundJdk = "FoundJdk"
Const FoundJre = "FoundJre"
Const FoundPath = "FoundPath"
Const FoundNone = "FoundNone"
Const strExit = "strExit"

Const strVarTypeSys = "System"
Const strVarTypeUsr = "User"
Const strWriteTypeAppend = "append"
Const strWriteTypeReplace = "replace"
Const strWriteTypeAddNew = "addnew"

Const javahomesys = "javahomesys"
Const javahomeusr = "javahomeusr"
Const jrehomesys = "jrehomesys"
Const jrehomeusr = "jrehomeusr"


Call StartJDKManager()



'###########################################################################


'-----------------------------------------------------
' **** STEP 1: READING WIN ENVIRONMENT VARIABLE **** '
'-----------------------------------------------------

Sub StartJDKManager()

Dim IsError

	Call ShowWelcomeBox()
	Call GetInstalledJDKJRE ()
	Call GetJavaHomeVars ()
	Call GetJavaPathVars ()
	
	Call PublishJDKJRE (dictInstalledJDKKRE, installedjava)
	Call PublishJDKJRE (dictHomeVars, homevars)
	Call PublishJDKJRE (dictPathVars, pathvars)
	
	If Not(dictInstalledJDKKRE.Exists("NoJDK")) And Not(dictInstalledJDKKRE.Exists("NoJRE")) Then
		strSelectedOption = ShowUserOptions(FoundJdkJre)	
		If (strSelectedOption <> strInvalid) Then
			Select Case strSelectedOption
			    Case strExit
			    	Call ExitApp()
			    Case strAllEnvVar
					For i = 1 To 1										
						If Not(ParseAndCallSetter(strJavaHome))Then 
							IsError = True
							Exit For
						End If
						If Not(ParseAndCallSetter(strJreHome)) Then
							IsError = True
							Exit For
						End If
						If Not(ParseAndCallSetter(strPathEnvVar)) Then
							IsError = True
							Exit For
						End If
					Next
					If IsError Then
						Call RestartCheck()
					End If
					Call RestartCheck()
			    Case Else
			    	ParseAndCallSetter(strSelectedOption)
					Call RestartCheck()
			End Select
		Else 
			Call RestartCheck()
		End If
	ElseIf Not(dictInstalledJDKKRE.Exists("NoJDK"))  And (dictInstalledJDKKRE.Exists("NoJRE")) Then
		strSelectedOption = ShowUserOptions(FoundJdk)
		If (strSelectedOption <> strInvalid) Then
			Select Case strSelectedOption
			    Case strExit
			    	Call ExitApp()
			    Case strJavaHomePathEnvVar
					For i = 1 To 1
						If Not(ParseAndCallSetter(strJavaHome)) Then
							IsError = True
							Exit For
						End If
						If Not(ParseAndCallSetter(strPathEnvVar)) Then
							IsError = True
							Exit For
						End If
					Next
					If IsError Then
						Call RestartCheck()
					End If	
					Call RestartCheck()				
			    Case Else
			    	ParseAndCallSetter(strSelectedOption)
			    	Call RestartCheck()
			End Select
		Else 
			Call RestartCheck()
		End If		
	ElseIf (dictInstalledJDKKRE.Exists("NoJDK"))  And Not(dictInstalledJDKKRE.Exists("NoJRE")) Then
			strSelectedOption = ShowUserOptions(FoundJre)
		If (strSelectedOption <> strInvalid) Then
				Select Case strSelectedOption
				    Case strExit
				    	Call ExitApp()
				    Case Else
				    	ParseAndCallSetter(strSelectedOption)
				    	Call RestartCheck()
				End Select
			Else 
				Call RestartCheck()
			End If				
	ElseIf (dictInstalledJDKKRE.Exists("NoJDK"))  And (dictInstalledJDKKRE.Exists("NoJRE")) Then
			Call ShowUserOptions(FoundNone)
			Call ExitApp()
	End If 

	'Call ExitApp()

End Sub



'###########################################################################


'-----------------------------------------
' **** READING WIN ENVIRONMENT VARIABLE **** '
'-----------------------------------------
Sub ReadEnvVar(envVarType, envVarName)

Dim arrStrVar, strVar
    Set objShell = CreateObject("WScript.Shell")
    Set objEnv = objShell.Environment(envVarType)
    WScript.Echo objEnv(envVarName)
    arrStrVar = Split(objEnv(envVarName), ";")

    For Each strVar In arrStrVar
        WScript.Echo strVar
    Next

Set objEnv = Nothing
Set objShell = Nothing

End Sub


'-------------------------------------------------------------------------
' **** STEP 7: WRITING WIN ENVIRONMENT VARIABLE **** '
'-------------------------------------------------------------------------

Sub WriteEnvVar(strEnvVarType, strEnvVarName, strEnvVarValue, strWriteType)

Dim strOldVarValue, strNewVarValue
Dim strVarType

Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment(strEnvVarType) ' envVarType= System or User

Select Case strEnvVarName
    Case strJavaHome
    	Select Case strWriteType
    	    Case strWriteTypeReplace
	    		strOldVarValue = objEnv(strEnvVarName)
	    		strNewVarValue = strEnvVarValue
	    		objEnv(strEnvVarName) = strNewVarValue
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "SETTING " & strEnvVarName & " TO " & strNewVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
			Case strWriteTypeAddNew
	    		objEnv(strEnvVarName) = strEnvVarValue
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "CREATING " & strEnvVarName & " AND SETTING TO " & strEnvVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
    	End Select
    	
    Case strJreHome
    	Select Case strWriteType
    	    Case strWriteTypeReplace
	    		strOldVarValue = objEnv(strEnvVarName)
	    		strNewVarValue = strEnvVarValue
	    		objEnv(strEnvVarName) = strNewVarValue
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "SETTING " & strEnvVarName & " TO " & strNewVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
			Case strWriteTypeAddNew
	    		objEnv(strEnvVarName) = strEnvVarValue
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "CREATING " & strEnvVarName & " AND SETTING TO " & strEnvVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
    	End Select    
    
    Case strPathEnvVar
    	Select Case strWriteType
    	    Case strWriteTypeReplace
	    		strOldVarValue = objEnv(strEnvVarName)
	    		strOldVarValue = PathExcludingJava(dictPathVars, arrPathTypes, strOldVarValue)
	    		strNewVarValue = strOldVarValue & ";" & strEnvVarValue
	    		objEnv(strEnvVarName) = strNewVarValue
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "SETTING " & strEnvVarName & " TO " & strNewVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
    	    Case strWriteTypeAppend
	    		strOldVarValue = objEnv(strEnvVarName)
	    		strNewVarValue = strOldVarValue & ";" & strEnvVarValue
	    		objEnv(strEnvVarName) = strNewVarValue
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "SETTING " & strEnvVarName & " TO " & strNewVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
			Case strWriteTypeAddNew
	    		WScript.StdOut.WriteBlankLines(1)
				WScript.StdOut.WriteLine "CREATING " & strEnvVarName & " AND SETTING TO " & strEnvVarValue & " ... DONE!"
				WScript.StdOut.WriteBlankLines(1)
				objEnv(strEnvVarName) = strEnvVarValue
    	End Select        
    	
End Select

Set objEnv = Nothing
Set objShell = Nothing

End Sub



'###########################################################################

Sub ReplaceStringInEnv()
	
	
	
	
End Sub


'###########################################################################


'----------------------------------------------------------
' **** STEP 2: LIST INSTALLED JDK/JRE From REGISTRY **** '
'----------------------------------------------------------

Function GetInstalledJDKJRE()

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."
strCallType = "CPanel"

arrRegLoc = Array("SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\","SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
'"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
'"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry2 = "InstallDate"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"
strEntry5 = "InstallLocation"

Dim strAppName, strAppVersion, strAppDate, strAppLoc
Dim arrJDK, arrJRE
Dim arrJDKTemp(), arrJRETemp()
Dim dictJDKJREOut

strJDKFound = 0
strJREFound = 0

Set dictJDKJREOut = CreateObject("Scripting.Dictionary")
Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")

For Each strKey In arrRegLoc
	If objReg.EnumKey (HKLM, strKey, arrSubkeys) = 0 Then 'success value is 0, meaning registry key exists
		For Each strSubkey In arrSubkeys
		  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
		  If intRet1 <> 0 Then
		    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
		  End If
		  If strValue1 <> "" Then
		    strAppName = strValue1
		    objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3
		    objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4
		    strAppVersion = intValue3 & "." & intValue4
		    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2
		    strAppDate = strValue2
		    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry5, strValue5
		    strAppLoc = strValue5
		    If (StrComp(Right(strAppLoc,1),"\") = 0) Then
		    	strAppLoc = Mid(strAppLoc,1,Len(strAppLoc)-1)
		    End if
		    If (InStr(strAppName, "Java") > 0) Then
		        Select Case IsJdkJreString(strAppName, strCallType)
		            Case JDK
		                arrJDK = ArrayFiller(arrJDKTemp, strAppName, strAppVersion, strAppDate, strAppLoc, strJDKFound)
		                strJDKFound = strJDKFound + 1
		            Case JRE
		                arrJRE = ArrayFiller(arrJRETemp, strAppName, strAppVersion, strAppDate, strAppLoc, strJREFound)
		                strJREFound = strJREFound + 1
		        End Select
		    End If
		  End If
		Next
	End If
Next

  If strJDKFound = 0 Then
    dictJDKJREOut.Add "NoJDK", "JDK Not Found"
  Else
    dictJDKJREOut.Add "JDK", arrJDK
  End If
  
  If strJREFound = 0 Then
    dictJDKJREOut.Add "NoJRE", "JRE Not Found"
  Else
    dictJDKJREOut.Add "JRE", arrJRE
  End If


Set objReg = Nothing
Set dictInstalledJDKKRE = dictJDKJREOut 
Set GetInstalledJDKJRE = dictJDKJREOut

End Function

'###########################################################################


'---------------------------------------------------
' **** EVALUATE IF JDK/JRE VALUES FOUND IN STRING **** '
'---------------------------------------------------
Function IsJdkJreString(strJdkJre, strCallType)

strJDK = "Development Kit"
strTypeJpath = "javapath"
strTypeCpath = "CPanel"

Select Case strCallType
    Case strTypeCpath
        If (InStr(1, strJdkJre, strJDK) <> 0) Then
            IsJdkJreString = JDK
        ElseIf (IsNumeric(Mid(strJdkJre, 6, 1))) Then
            IsJdkJreString = JRE
        Else
            IsJdkJreString = False
        End If
    Case strTypeJpath
        If (InStr(strJdkJre, strTypeJpath) <> 0) Then
            IsJdkJreString = True
        Else
            IsJdkJreString = False
        End If
End Select

End Function

'###########################################################################


'-----------------------------------------
' **** FILL JDK/JRE ARRAY **** '
'-----------------------------------------
Function ArrayFiller(ByRef arrJDKJRE, strAppName, strAppVersion, strAppDate, strAppLoc, strJDKJREFound)

ReDim Preserve arrJDKJRE(3, strJDKJREFound)
arrJDKJRE(0, strJDKJREFound) = strAppName
arrJDKJRE(1, strJDKJREFound) = strAppVersion
arrJDKJRE(2, strJDKJREFound) = strAppDate
arrJDKJRE(3, strJDKJREFound) = strAppLoc
ArrayFiller = arrJDKJRE

End Function

'###########################################################################


'-----------------------------------------
' **** STEP 5: PUBLISH OUTPUT **** '
'-----------------------------------------

Sub PublishJDKJRE (oDataDict, strDictType)
Dim strTemp

Select Case strDictType
    
    Case installedjava 	
    	WScript.StdOut.WriteBlankLines(1) 	
    	If oDataDict.Exists(JDK) Then
    		WScript.StdOut.WriteLine(vbCrLf & "JDK INSTALLATIONS FOUND ON SYSTEM :-" & vbCrLf & "----------------------------------")
    		WScript.StdOut.WriteBlankLines(1)
    		Call ArrayIterator(oDataDict("JDK"), 3)
    		WScript.StdOut.WriteBlankLines(1)
    	Else
    		WScript.StdOut.WriteLine(vbCrLf & "JDK INSTALLATIONS FOUND ON SYSTEM :-" & vbCrLf & "----------------------------------")
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine ("NO JDK INSTALLATIONS FOUND IN REGISTRY AND CONTROL PANEL..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If
    	If oDataDict.Exists(JRE) Then
    		WScript.StdOut.WriteLine(vbCrLf & "JRE INSTALLATIONS FOUND ON SYSTEM :-" & vbCrLf & "---------------------------------")
    		WScript.StdOut.WriteBlankLines(1)
    		Call ArrayIterator(oDataDict("JRE"), 3)
    		WScript.StdOut.WriteBlankLines(1)
    	Else
    		WScript.StdOut.WriteLine(vbCrLf & "JRE INSTALLATIONS FOUND ON SYSTEM :-" & vbCrLf & "---------------------------------")
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine ("NO JRE INSTALLATIONS FOUND IN REGISTRY AND CONTROL PANEL ..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If
	Case homevars
		    WScript.StdOut.WriteBlankLines(2)
    		WScript.StdOut.WriteLine(vbCrLf & "SYSTEM VARIABLES CURRENTLY SET:-" & vbCrLf & "------------------------------")
    		WScript.StdOut.WriteBlankLines(1)    		
    	If oDataDict.Exists(javahomesys) Then
    		strTemp = oDataDict.Item(javahomesys)
    		WScript.StdOut.WriteLine "JAVA_HOME = " & strTemp
    	Else
    		WScript.StdOut.WriteLine ("JAVA_HOME = CURRENTLY NOT SET ..!")
    	End If
    	If oDataDict.Exists(jrehomesys) Then
    	    strTemp = oDataDict.Item(jrehomesys)
    		WScript.StdOut.WriteLine "JRE_HOME = " & strTemp
    	Else
    		WScript.StdOut.WriteLine ("JRE_HOME = CURRENTLY NOT SET ..!")
    	End If
    		WScript.StdOut.WriteBlankLines(2)
    		WScript.StdOut.WriteLine(vbCrLf & "USER VARIABLES CURRENTLY SET :-" & vbCrLf & "----------------------------")
    		WScript.StdOut.WriteBlankLines(1)    		
    	If oDataDict.Exists(javahomeusr) Then
    		strTemp = oDataDict.Item(javahomeusr)
    		WScript.StdOut.WriteLine "JAVA_HOME = " & strTemp 
    	Else
    		WScript.StdOut.WriteLine ("JAVA_HOME = CURRENTLY NOT SET ..!")
    	End If
    	If oDataDict.Exists(jrehomeusr) Then
    		strTemp = oDataDict.Item(jrehomeusr)
    	   WScript.StdOut.WriteLine "JAVA_HOME = " & strTemp 
    	Else
    		WScript.StdOut.WriteLine ("JRE_HOME = CURRENTLY NOT SET ..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If    	
    	WScript.StdOut.WriteBlankLines(1)
	Case pathvars
			WScript.StdOut.WriteBlankLines(2)
    		WScript.StdOut.WriteLine(vbCrLf & "PATH VARIABLES CURRENTLY SET :-" & vbCrLf & "----------------------------")	
    		WScript.StdOut.WriteBlankLines(1)
		If oDataDict.Exists("NoPathVars") Then
			WScript.StdOut.WriteLine "PATH VARIABLE CURRENTLY NOT SET ..!"
    	ElseIf oDataDict.Exists("NoJavaPathVars") Then
    		WScript.StdOut.WriteLine "NO 'JAVA' PATH CURRENTLY SET IN 'PATH' VARIABLE ..!"
		ElseIf oDataDict.Exists("javapath") Then
			strTemp = oDataDict.Item("javapath")
    		WScript.StdOut.WriteLine strTemp
    	ElseIf oDataDict.Exists("%JAVA_HOME%\bin") Then
    		strTemp = oDataDict.Item("%JAVA_HOME%\bin")
    		WScript.StdOut.WriteLine strTemp    	
    	ElseIf oDataDict.Exists("jdk") Then
    		strTemp = oDataDict.Item("jdk")
    		WScript.StdOut.WriteLine strTemp
    	ElseIf oDataDict.Exists("jre") Then
    		strTemp = oDataDict.Item("jre")
    		WScript.StdOut.WriteLine strTemp
    	End If
		WScript.StdOut.WriteBlankLines(2)
End Select


End Sub

'###########################################################################


Sub ArrayIterator(arrDataObj, strDimension)

Dim counter, i
Select Case strDimension
    Case 1
            WScript.StdOut.WriteLine (arrDataObj(0) & " = " & arrDataObj(1))
    Case 3
        For counter = 0 To UBound(arrDataObj, 2)
			WScript.StdOut.WriteLine ("[" & counter+1 & "] " & arrDataObj(0, counter)) 
            WScript.StdOut.WriteLine ("    " & "Version: " & arrDataObj(1, counter))
            WScript.StdOut.WriteLine ("    " & "Install Date: " & arrDataObj(2, counter))
            WScript.StdOut.WriteLine ("    " & "Location: " & arrDataObj(3, counter))
        Next
End Select

End Sub

'###########################################################################


'--------------------------------------------------
' **** STEP 3: LIST JDK/JRE ENV HOME VARAIABLES **** '
'--------------------------------------------------

Function GetJavaHomeVars()
'Using WMI retrieves both USER and SYSTEM variable together, you cannot pick and choose

strCntJDKSys = 0
strCntJRESys = 0
strCntJDKUsr = 0
strCntJREUsr = 0

strComputer = "."

Set dictJavaVarOut = CreateObject("scripting.dictionary")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")

For Each objItem In colItems

    If (StrComp(objItem.Name, strJavaHome) = 0) Then
        Select Case (objItem.SystemVariable)
            Case "True"
                dictJavaVarOut.Add "javahomesys", objItem.VariableValue
				strCntJDKSys = strCntJDKSys + 1
            Case "False"
                dictJavaVarOut.Add "javahomeusr", objItem.VariableValue
                strCntJDKUsr = strCntJDKUsr + 1
        End Select
    ElseIf (StrComp(objItem.Name, strJreHome) = 0) Then
        Select Case (objItem.SystemVariable)
            Case "True"
                dictJavaVarOut.Add "jrehomesys", objItem.VariableValue
                strCntJRESys = strCntJDKSys + 1
            Case "False"
                dictJavaVarOut.Add "jrehomeusr", objItem.VariableValue
                strCntJREUsr = strCntJDKUsr + 1
        End Select
    End If
Next

If Not(dictJavaVarOut.Exists("javahomesys")) Then
	dictJavaVarOut.Add "nojavahomesys", "no java home in system vars"
End If
If Not(dictJavaVarOut.Exists("javahomeusr")) Then
	dictJavaVarOut.Add "nojavahomeusr", "no java home in user vars"
End If
If Not(dictJavaVarOut.Exists("jrehomesys")) Then
	dictJavaVarOut.Add "nojrehomesys", "no jre home in system vars"
End If
If Not(dictJavaVarOut.Exists("jrehomeusr")) Then
	dictJavaVarOut.Add "nojrehomeusr", "no jre home in user vars"
End If


If (dictJavaVarOut.Count) > 0 Then
	Set dictHomeVars = dictJavaVarOut
	Set GetJavaHomeVars = dictJavaVarOut
Else
	' ELABORATE THIS. ITS USEFUL IN CASE NONE OF THE VARIABLE ARE SET OR DISCOVERED
	dictJavaVarOut.Add "NoVars", "No Vars Found" 
	GetJavaHomeVars = False
End If


Set colItems = Nothing
Set objWMIService = Nothing

End Function


'###########################################################################

'--------------------------------------------------
' **** STEP 4: LIST PATH ENV VARAIABLES **** '
'--------------------------------------------------
Function GetJavaPathVars()

Dim strExtPath, dictJavaPathOut
strComputer = "."

strCountFound = 0
strPathEnvSet = 0

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")
Set dictJavaPathOut = CreateObject("scripting.dictionary")

For Each objItem In colItems    
    If (StrComp(objItem.Name, strPathEnvVar) = 0) Then
    	strPathEnvSet = strPathEnvSet + 1
        For Each strExp In arrPathTypes
        If (InStr(objItem.VariableValue, strExp) <> 0) Then
            strExtPath = ExtractPathValue(objItem.VariableValue, strExp)
            If strExtPath <> False Then
                dictJavaPathOut.Add strExp, strExtPath
                strCountFound = strCountFound + 1
            End If
        End If
        Next
    End If
Next

If (strPathEnvSet = 0) Then
	dictJavaPathOut.Add NoPathVars, "No Path Vars Found" 
	Set dictPathVars = dictJavaPathOut
	Set GetJavaPathVars = dictJavaPathOut
ElseIf (strCountFound = 0) Then
	dictJavaPathOut.Add NoJavaPathVars, "No Path Vars Found" 
		Set dictPathVars = dictJavaPathOut
	Set GetJavaPathVars = dictJavaPathOut
Else 
	Set dictPathVars = dictJavaPathOut
	Set GetJavaPathVars = dictJavaPathOut
End If

'Call PublishJDKJRE (dictJavaPathOut, "pathvars") '**** DELETE THIS ****

Set colItems = Nothing
Set objWMIService = Nothing

End Function

'###########################################################################

'--------------------------------------------------
' **** EXTRACT VALUE FROM SYS VARIABLES **** '
'--------------------------------------------------
Function ExtractPathValue(strFullPathValue, strValueType)

Dim strJavaPath, strEndPos, cntExtract
Dim iTemp

strMatchPos = InStr(strFullPathValue, strValueType) + (Len(strValueType) - 1)

For iTemp = strMatchPos To Len(strFullPathValue)
    If (Mid(strFullPathValue, iTemp, 1) = ";") Then
        Exit For
    End If
Next

strEndPos = (iTemp - 1)
cntExtract = strEndPos
strNew = Mid(strFullPathValue, 1, strEndPos)

For i = 1 To strEndPos
    If (Mid(strNew, cntExtract, 1) = ";") Then
        Exit For
    Else
        strJavaPath = strJavaPath & Mid(strNew, cntExtract, 1)
        cntExtract = cntExtract - 1
    End If
Next

ExtractPathValue = StrReverse(strJavaPath)

End Function

'###########################################################################

'-------------------------------------------------------------------------
' **** STEP 8: REMOVE PREVIOUS JAVA VALUES FROM PREVIOUS PATH VAR **** '
'-------------------------------------------------------------------------

Function PathExcludingJava(dictPaths, arrPathStrTypes, strCurPathValue)

Dim strInterimPathArray1, strInterimPathArray2
Dim strToReplace

strInterimPathArray1 = Split(strCurPathValue, ";", -1, 1)

    For Each strExp In arrPathStrTypes
        If dictPaths.Exists(strExp) Then
            strToReplace = dictPaths.Item(strExp)
            strInterimPathArray2 = Filter(strInterimPathArray1, strToReplace, False)
            strInterimPathArray1 = strInterimPathArray2
        End If
    Next
        
    PathExcludingJava = Join(strInterimPathArray1, ";")
     
End Function


'###########################################################################


'-------------------------------------------------------------------------
' **** STEP 6: PARSE USER INPUT AND RE-DIRECT TO SPECIFIC SETTER **** '
'-------------------------------------------------------------------------

Function ParseAndCallSetter(strOptionInput) 'javahome

Dim arrJDK, arrJRE
Dim strSelectedJDK, strValue, strVarType

Select Case strOptionInput
    Case strJavaHome
    	arrJDK = dictInstalledJDKKRE.Item(JDK)
    	strSelectedJDK = ShowJDKOptions(arrJDK, strJavaHome,"")
    	If (strSelectedJDK <> strInvalid) Then
	    	strValue = arrJDK(3,strSelectedJDK)
	    	If dictHomeVars.Exists(javahomesys) Then
	    		WriteEnvVar strVarTypeSys,strJavaHome,strValue, strWriteTypeReplace
	    	Else
	    		WriteEnvVar strVarTypeSys,strJavaHome,strValue, strWriteTypeAddNew
	    	End If
			WScript.StdOut.WriteBlankLines(1)
			WScript.StdOut.WriteLine "VERIFYING, IF CHANGES HAVE PERSISTED ... "
			WScript.StdOut.WriteBlankLines(1)
			WScript.StdOut.WriteLine "CURRENT STATUS POST-CHANGES BELOW ... "
			Call PublishJDKJRE (GetJavaHomeVars(), homevars)
			ParseAndCallSetter = True
		Else 
			ParseAndCallSetter = False
		End If
    Case strJreHome
    	arrJRE = dictInstalledJDKKRE.Item(JRE)
    	strSelectedJDK = ShowJDKOptions(arrJRE, strJreHome,"")
    	If (strSelectedJDK <> strInvalid) Then
	    	strValue = arrJRE(3,strSelectedJDK)
	    	If dictHomeVars.Exists(jrehomesys) Then
	    		WriteEnvVar strVarTypeSys,strJreHome,strValue, strWriteTypeReplace
	    	Else
	    		WriteEnvVar strVarTypeSys,strJreHome,strValue, strWriteTypeAddNew
	    	End If
	    	WScript.StdOut.WriteBlankLines(1)
			WScript.StdOut.WriteLine "VERIFYING, IF CHANGES HAVE PERSISTED ... "
			WScript.StdOut.WriteBlankLines(1)
			WScript.StdOut.WriteLine "CURRENT STATUS POST-CHANGES BELOW ... "
			Call PublishJDKJRE (GetJavaHomeVars(), homevars)
			ParseAndCallSetter = True
		Else 
			ParseAndCallSetter = False
		End If
    Case strPathEnvVar
    	arrJDK = dictInstalledJDKKRE.Item(JDK)
    	strSelectedJDK = ShowJDKOptions(arrJDK, strJavaHome, strPathEnvVar)
    	If (strSelectedJDK <> strInvalid) Then
	    	strValue = arrJDK(3,strSelectedJDK)
	    	strValue = strValue & "\bin"
	    	If dictPathVars.Exists(NoJavaPathVars) Then
	    		WriteEnvVar strVarTypeSys,strPathEnvVar,strValue, strWriteTypeAppend
	    	ElseIf dictPathVars.Exists(NoPathVars) Then
	    		WriteEnvVar strVarTypeSys,strPathEnvVar,strValue, strWriteTypeAddNew
	    	Else 
	    		WriteEnvVar strVarTypeSys,strPathEnvVar,strValue, strWriteTypeReplace
	    	End If
	    	WScript.StdOut.WriteBlankLines(1)
			WScript.StdOut.WriteLine "VERIFYING, IF CHANGES HAVE PERSISTED ... "
			WScript.StdOut.WriteBlankLines(1)
			WScript.StdOut.WriteLine "CURRENT STATUS POST-CHANGES BELOW ... "
			Call PublishJDKJRE (GetJavaPathVars(), pathvars)
			ParseAndCallSetter = True
		Else 
			ParseAndCallSetter = False
		End If			
    Case Else
        WScript.StdOut.WriteBlankLines (1)
        WScript.StdOut.WriteLine "ERROR! INVALID CHOICE !"
End Select


End Function


'###########################################################################

Public Sub ShowWelcomeBox()

WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & vbTab & VBTab & "    " & "WIN-JDK-MANAGER v1.0"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "VBScript (WMI,WScript) Utility. View all installed JDK/JRE"
WScript.StdOut.WriteLine vbTab & " " & "versions [32bit/64bit].Easily view and re-point Env Vars"
WScript.StdOut.WriteLine VBTab & "    " & "Platform: Win7/8 | Pre-Req: Script/Admin Privilege"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "   " & "Updated: Sept 2019 | Tushar Sharma | www.testoxide.com"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(2)

End Sub

'###########################################################################

Public Function ConsoleInput()
	ConsoleInput = WScript.StdIn.ReadLine
End Function

'###########################################################################


Sub ExitApp()
	 WScript.StdOut.WriteBlankLines(1)
	 WScript.StdOut.WriteLine "Press 'Enter' key to exit ..."
	 ConsoleInput()
	 WScript.Quit
End Sub

'###########################################################################


Function ShowUserOptions(strShowOption)

Dim strUsrInput, strAvailableOptions

WScript.StdOut.WriteBlankLines(2)
WScript.StdOut.Write "             ~~~~~~~~~~   STARTING <INTERACTIVE MODE>   ~~~~~~~~~~              "
WScript.StdOut.WriteBlankLines(2)

Select Case strShowOption
    Case FoundJdkJre
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "CHOOSE ONE OF THE BELOW AVAILABLE OPTIONS [Eg. Choose 1 for Setting JAVA_HOME]"
		WScript.StdOut.WriteBlankLines(2)
		WScript.StdOut.WriteLine "[1] Set JAVA_HOME (SYSTEM) Env Variable Only ?" & vbCrLf
		WScript.StdOut.WriteLine "[2] Set JRE_HOME (SYSTEM) Env Variable Only ?" & vbCrLf
		WScript.StdOut.WriteLine "[3] Set PATH (SYSTEM) Env Variable Only ?" & vbCrLf
		WScript.StdOut.WriteLine "[4] Set all above i.e. JAVA_HOME, JRE_HOME and PATH ?" & vbCrLf
		WScript.StdOut.WriteLine "[5] Cancel and Exit ?"
		WScript.StdOut.WriteBlankLines(2)
		WScript.StdOut.WriteLine "Tip: Type a bullet number from above and hit Enter."
		WScript.StdOut.WriteBlankLines(1)
	Case FoundJdk
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "CHOOSE ONE OF THE BELOW AVAILABLE OPTIONS? [Eg. Choose 1 for Setting JAVA_HOME]"
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "1. Set JAVA_HOME (SYSTEM) Env Variable Only ?"
		WScript.StdOut.WriteLine "2. Set PATH (SYSTEM) Env Variable Only ?"
		WScript.StdOut.WriteLine "3. Set both above i.e. JAVA_HOME and PATH ?"
		WScript.StdOut.WriteLine "4. Cancel and Exit ?"
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "Tip: Type a bullet number from above and hit Enter."
		WScript.StdOut.WriteBlankLines(1)
	Case FoundJre
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "CHOOSE ONE OF THE BELOW AVAILABLE OPTIONS? [Eg. Choose 1 for Setting JRE_HOME]"
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "1. Set JRE_HOME (SYSTEM) Env Variable Only ?"
		WScript.StdOut.WriteLine "2. Cancel and Exit ?"
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "Tip: Type a bullet number from above and hit Enter."
		WScript.StdOut.WriteBlankLines(1)
	Case FoundNone
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "YOU DON NOT HAVE ANY OPTION TO SET ENV VARIABLES, AS NO JAVA FOUND ON YOUR SYSTEM"
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "Tip: INSTALL JAVA AND RE-RUN THIS SCRIPT UTILITY"
		WScript.StdOut.WriteBlankLines(1)
End Select

'COLLECT USER INPUT
strUsrInput = ConsoleInput()

'If (ValidateInput(strUsrInput,5)) Then 
	Select Case strShowOption
	    Case FoundJdkJre
			If (ValidateInput(strUsrInput,5)) Then
				If strUsrInput = "1" Then
					ShowUserOptions = strJavaHome
				ElseIf (strUsrInput = "2") Then
					ShowUserOptions = strJreHome
				ElseIf (strUsrInput = "3") Then
					ShowUserOptions = strPathEnvVar
				ElseIf (strUsrInput = "4") Then
					ShowUserOptions = strAllEnvVar
				ElseIf (strUsrInput = "5") Then
					ShowUserOptions = strExit
				End If
			Else
				ShowUserOptions = strInvalid
			End If
		Case FoundJdk
			If (ValidateInput(strUsrInput,4)) Then
				If strUsrInput = "1" Then
					ShowUserOptions = strJavaHome
				ElseIf (strUsrInput = "2") Then
					ShowUserOptions = strPathEnvVar
				ElseIf (strUsrInput = "3") Then
					ShowUserOptions = strJavaHomePathEnvVar
				ElseIf (strUsrInput = "4") Then
					ShowUserOptions = strExit
				End If
			Else
				ShowUserOptions = strInvalid			
			End If
		Case FoundJre
			If (ValidateInput(strUsrInput,2)) Then
				If strUsrInput = "1" Then
					ShowUserOptions = strJreHome
				ElseIf (strUsrInput = "2") Then
					ShowUserOptions = strExit			
				End If
			Else
				ShowUserOptions = strInvalid				
			End If
		Case FoundNone
		
	End Select
	
'Else
'	ShowUserOptions = False
'End If


End Function

'###########################################################################

Function ShowJDKOptions(arrOJDKbj, strVarTypeSelect, IsPath)
Dim strUsrInput

Select Case strVarTypeSelect
    Case strJavaHome
    	If IsPath <> "" Then
			WScript.StdOut.WriteBlankLines(2)
			WScript.StdOut.WriteLine "SETTING PATH : CHOOSE ONE OF THE AVAILABLE JDKs"
			WScript.StdOut.WriteLine "-------------"	
			WScript.StdOut.WriteBlankLines(1)
		Else
			WScript.StdOut.WriteBlankLines(2)
			WScript.StdOut.WriteLine "SETTING JAVA_HOME : CHOOSE ONE OF THE AVAILABLE JDKs"
			WScript.StdOut.WriteLine "-----------------"	
			WScript.StdOut.WriteBlankLines(1)			
    	End If
		For counter = 0 To UBound(arrOJDKbj, 2)
			WScript.StdOut.WriteLine counter + 1 & ". " & arrOJDKbj(0, counter)
		Next
    Case strJreHome
		WScript.StdOut.WriteBlankLines(2)
		WScript.StdOut.WriteLine "SETTING JRE_HOME : CHOOSE ONE OF THE AVAILABLE JREs"
		WScript.StdOut.WriteLine "-----------------"
		WScript.StdOut.WriteBlankLines(1)
		For counter = 0 To UBound(arrOJDKbj, 2)
			WScript.StdOut.WriteLine counter + 1 & ". " & arrOJDKbj(0, counter)
		Next
End Select

	WScript.StdOut.WriteBlankLines(1)
	WScript.StdOut.WriteLine "Tip: Type a bullet number from above and hit Enter."
	WScript.StdOut.WriteBlankLines(1)
	
	strUsrInput = ConsoleInput()
	
	If (ValidateInput(strUsrInput, UBound(arrOJDKbj, 2) + 1)) Then
		ShowJDKOptions = strUsrInput-1	
	Else
		ShowJDKOptions = strInvalid
	End If

End Function


'###########################################################################

Function ValidateInput (strArgsIn, iMaxVal)

Dim strValidInput, strArg, strFound
strFound = False

strValidNumIn = Array("1","2","3","4","5","6")
strValidStrIn = Array("Y","N","YES","NO")

If (IsNumeric(strArgsIn)) Then
	If (CInt(strArgsIn) <= CInt(iMaxVal)) And (strArgsIn <> 0) Then
		For Each strArg In strValidNumIn
			If (StrComp(strArg, strArgsIn) = 0) Then
				strFound = True
				Exit For
			End If
		Next
	Else 
		strFound = False
	End If
Else
	For Each strArg In strValidStrIn
		If (StrComp(UCase(strArg), strArgsIn) = 0) Then
			strFound = True
			Exit For
		End If
	Next
End If
	
	If Not(strFound) Then
		ValidateInput = False
		WScript.StdOut.WriteBlankLines(1)
		WScript.StdOut.WriteLine "ERROR : INVALID INPUT! PLEASE RE-LOAD THE PROGRAM AND TRY AGAIN."
	Else 
		ValidateInput = True
	End If

End Function 


'###########################################################################

Function CheckArrayData (arrInput)

  IsArrayDimmed = False
  If IsArray(arr) Then
    On Error Resume Next
    Dim ub : ub = UBound(arr)
    If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
  End If  
  
'OR  
  
  iRet = True

    If IsArray(myArray) Then
        i = 0
        For Each e In myArray
            If Not IsEmpty(e) And Len(e)>0 Then
                i = i +1
            End If
        Next
        If i>0 Then
            iRet = False
        End If
    End If
    wIsArrayEmpty = iRet

End Function

'###########################################################################


Function IsReloadExit ()

WScript.StdOut.WriteBlankLines(2)
WScript.StdOut.WriteLine "RE-LOAD THE PROGRAM OR EXIT (y=Reload / n=Exit) ?"
strResponse = UCase(ConsoleInput())

If ValidateInput(strResponse,"") Then
	Select Case strResponse
	    Case "Y"
	    	IsReloadExit = True
	    Case "N"
	    	IsReloadExit = False
	End Select
Else
	WScript.StdOut.WriteLine "INVALID CHOICE!"
	ExitApp()
End If

End Function


'###########################################################################


Sub RestartCheck ()
	If IsReloadExit() Then
		Call StartJDKManager()
	Else
		ExitApp()
	End If
End Sub


'###########################################################################


'===========================================================================================================
'===========================================================================================================
