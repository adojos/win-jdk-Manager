
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\',
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'

'If WScript.Arguments.length = 0 Then
'   Set objShell = CreateObject("Shell.Application")
'   objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 3
'      WScript.Quit
'End If   


Call ShowWelcomeBox()

Call GetInstalledJDKJRE ()
Call GetJavaHomeVars ()
Call GetJavaPathVars()
Call SelectFunction ()


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

'-----------------------------------------
' **** WRITING WIN ENVIRONMENT VARIABLE **** '
'-----------------------------------------
Sub WriteEnvVar(envVarType, envVarName, envVarValue)

Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment(envVarType)
objEnv(envVarName) = envVarValue

Set objEnv = Nothing
Set objShell = Nothing

End Sub

'###########################################################################


'--------------------------------------------------
' **** LIST INSTALLED JDK/JRE From REGISTRY **** '
'--------------------------------------------------
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
		    If (InStr(strAppName, "Java") > 0) Then
		        Select Case IsJdkJreString(strAppName, strCallType)
		            Case "JDK"
		                arrJDK = ArrayFiller(arrJDKTemp, strAppName, strAppVersion, strAppDate, strAppLoc, strJDKFound)
		                strJDKFound = strJDKFound + 1
		            Case "JRE"
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

Call PublishJDKJRE (dictJDKJREOut, "installedjava") ' **** DELETE THIS ****

Set objReg = Nothing
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
            IsJdkJreString = "JDK"
        ElseIf (IsNumeric(Mid(strJdkJre, 6, 1))) Then
            IsJdkJreString = "JRE"
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
' **** PUBLISH OUTPUT **** '
'-----------------------------------------
Sub PublishJDKJRE (oDataDict, strDictType)

Select Case strDictType
    
    Case "installedjava"  	
    	WScript.StdOut.WriteBlankLines(1)
    	WScript.StdOut.WriteLine("==================================" & vbCrLf & "JAVA INSTALLATIONS FOUND ON SYSTEM" & vbCrLf & "==================================")
    	WScript.StdOut.WriteBlankLines(1)
    	
    	If oDataDict.Exists("JDK") Then
    		WScript.StdOut.WriteLine(vbCrLf & "JDK INSTALLATIONS :-" & vbCrLf & "------------------")
    		Call ArrayIterator(oDataDict("JDK"), 3)
    		WScript.StdOut.WriteBlankLines(1)
    	Else
    		WScript.StdOut.WriteLine(vbCrLf & "JDK INSTALLATIONS :-" & vbCrLf & "------------------")
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine ("NO JDK INSTALLATIONS FOUND IN REGISTRY AND CONTROL PANEL..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If
    	If oDataDict.Exists("JRE") Then
    		WScript.StdOut.WriteLine(vbCrLf & "JRE INSTALLATIONS :-" & vbCrLf & "-----------------")
    		Call ArrayIterator(oDataDict("JRE"), 3)
    		WScript.StdOut.WriteBlankLines(1)
    	Else
    		WScript.StdOut.WriteLine(vbCrLf & "JRE INSTALLATIONS :-" & vbCrLf & "-----------------")
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine ("NO JRE INSTALLATIONS FOUND IN REGISTRY AND CONTROL PANEL ..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If
	Case "homevars"
		    WScript.StdOut.WriteBlankLines(2)
		    WScript.StdOut.WriteLine("==================================" & vbCrLf & "ENVIRONMENT VARIABLES CURRENTLY SET" & vbCrLf & "==================================")
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine(vbCrLf & "SYSTEM VARIABLES :-" & vbCrLf & "----------------")
    	If oDataDict.Exists("javahomesys") Then
    		Call ArrayIterator(oDataDict("javahomesys"), 1)
    	Else
    		WScript.StdOut.WriteLine ("JAVA_HOME = CURRENTLY NOT SET ..!")
    	End If
    	If oDataDict.Exists("jrehomesys") Then
    	   Call ArrayIterator(oDataDict("jrehomesys"), 1)
    	Else
    		WScript.StdOut.WriteLine ("JRE_HOME = CURRENTLY NOT SET ..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine(vbCrLf & "USER VARIABLES :-" & vbCrLf & "--------------")
    	If oDataDict.Exists("javahomeusr") Then
    		Call ArrayIterator(oDataDict("javahomeusr"), 1)
    	Else
    		WScript.StdOut.WriteLine ("JAVA_HOME = CURRENTLY NOT SET ..!")
    	End If
    	If oDataDict.Exists("jrehomeusr") Then
    	   Call ArrayIterator(oDataDict("jrehomeusr"), 1)
    	Else
    		WScript.StdOut.WriteLine ("JRE_HOME = CURRENTLY NOT SET ..!")
    		WScript.StdOut.WriteBlankLines(1)
    	End If    	
	Case "pathvars"
    		WScript.StdOut.WriteBlankLines(1)
    		WScript.StdOut.WriteLine(vbCrLf & "PATH VARIABLES :-" & vbCrLf & "----------------")	
		If oDataDict.Exists("NoPathVars") Then
    		WScript.StdOut.WriteLine "NO JAVA PATH CURRENTLY SET IN 'PATH' VARIABLE ..!"
    	End If
		If oDataDict.Exists("javapath") Then
    		WScript.StdOut.WriteLine (oDataDict("javapath"))
    	End If
    	If oDataDict.Exists("%JAVA_HOME%\bin") Then
    		WScript.StdOut.WriteLine (oDataDict("%JAVA_HOME%\bin"))
    	End If
    	If oDataDict.Exists("jdk") Then
    		WScript.StdOut.WriteLine (oDataDict("jdk"))
    	End If
    	If oDataDict.Exists("jre") Then
    		WScript.StdOut.WriteLine (oDataDict("jre"))
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
' **** LIST JDK/JRE ENV HOME VARAIABLES **** '
'--------------------------------------------------
Function GetJavaHomeVars()
'Using WMI retrieves both USER and SYSTEM variable together, you cannot pick and choose

Dim arrJDKSys, arrJRESys
strCntJDKSys = 0
strCntJRESys = 0

Dim arrJDKUsr, arrJREUsr
strCntJDKUsr = 0
strCntJREUsr = 0

strComputer = "."

Set dictJavaVarOut = CreateObject("scripting.dictionary")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")

For Each objItem In colItems

    If (StrComp(objItem.Name, "JAVA_HOME") = 0) Then
        Select Case (objItem.SystemVariable)
            Case "True"
                arrJDKSys = Array(objItem.Name, objItem.VariableValue)
                dictJavaVarOut.Add "javahomesys", arrJDKSys
				strCntJDKSys = strCntJDKSys + 1
            Case "False"
                arrJDKUsr = Array(objItem.Name, objItem.VariableValue)
                dictJavaVarOut.Add "javahomeusr", arrJDKUsr
                strCntJDKUsr = strCntJDKUsr + 1
        End Select
    ElseIf (StrComp(objItem.Name, "JRE_HOME") = 0) Then
        Select Case (objItem.SystemVariable)
            Case "True"
                arrJRESys = Array(objItem.Name, objItem.VariableValue)
                dictJavaVarOut.Add "jrehomesys", arrJRESys
                strCntJRESys = strCntJDKSys + 1
            Case "False"
                arrJREUsr = Array(objItem.Name, objItem.VariableValue)
                dictJavaVarOut.Add "jrehomeusr", arrJREUsr
                strCntJREUsr = strCntJDKUsr + 1
        End Select
    End If
Next

If (dictJavaVarOut.Count) > 0 Then
	Set GetJavaHomeVars = dictJavaVarOut
Else
	dictJavaVarOut.Add "NoVars", "No Vars Found" 
	GetJavaHomeVars = False
End If

Call PublishJDKJRE (dictJavaVarOut, "homevars") '**** DELETE THIS ****

Set colItems = Nothing
Set objWMIService = Nothing

End Function


'###########################################################################

'--------------------------------------------------
' **** LIST PATH ENV VARAIABLES **** '
'--------------------------------------------------
Function GetJavaPathVars()

arrPathType = Array("javapath", "%JAVA_HOME%\bin", "jdk", "jre")
Dim strExtPath, dictJavaPathOut
strComputer = "."


Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")
Set dictJavaPathOut = CreateObject("scripting.dictionary")

For Each objItem In colItems
    If (StrComp(objItem.Name, "Path") = 0) Then
        For Each strExp In arrPathType
        If (InStr(objItem.VariableValue, strExp) <> 0) Then
            strExtPath = ExtractPathValue(objItem.VariableValue, strExp)
            If strExtPath <> False Then
                dictJavaPathOut.Add strExp, strExtPath
            End If
        End If
        Next
    End If
Next

If (dictJavaPathOut.Count) > 0 Then
	Set GetJavaPathVars = dictJavaPathOut
Else 
	dictJavaPathOut.Add "NoPathVars", "No Path Vars Found" 
	GetJavaPathVars = False
End If

Call PublishJDKJRE (dictJavaPathOut, "pathvars") '**** DELETE THIS ****

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

Sub GetJavaWmiJSoftReg()

End Sub

'###########################################################################

Public Sub ShowWelcomeBox()

WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & vbTab & VBTab & "    " & "win-JDK-Manager v1.0"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "VBScript (WMI,WScript) Utility. View all installed JDK/JRE"
WScript.StdOut.WriteLine vbTab & " " & "versions [32bit/64bit].Easily view and re-point Env Vars"
WScript.StdOut.WriteLine VBTab & "    " & "Platform: Win7/8 | Pre-Req: Script/Admin Privilege"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "  " & "Last Updated: Wed, 15 May 2019 | Author: Tushar Sharma"
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


Public Function SelectFunction()
Dim strFun

WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "SELECT FUNCTION TO PERFORM? [Eg. Type 1 for Setting JAVA_HOME]"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "1. Set JAVA_HOME Env Variable"
WScript.StdOut.WriteLine "2. Set JRE_HOME Env Variable"
WScript.StdOut.WriteLine "3. Set PATH Env Variable"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "Tip: Type a number from above and hit Enter."
WScript.StdOut.WriteBlankLines(1)

strFun = ConsoleInput()
SelectFunction = strFun

End Function


'==================================================================
'==================================================================
'==================================================================

'-----------------------------------------
' **** LIST ALL INSTALLED SOFTWARES **** '
'-----------------------------------------
Sub ListAllInstalledApps()

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry2 = "InstallDate"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"

arrJDK = ""
arrJRE = ""

Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
 
objReg.EnumKey HKLM, strKey, arrSubkeys
'WScript.Echo "Installed Applications" & vbCrLf

For Each strSubkey In arrSubkeys
  
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
   strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry1b, strValue1
  End If
  If strValue1 <> "" Then
    'WScript.Echo VbCrLf & "Application Name: " & strValue1
    'WScript.Echo vbCrLf & strValue1
  End If
  
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry3, intValue3
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry4, intValue4
  If intValue3 <> "" Then
     'WScript.Echo "Version: " & intValue3 & "." & intValue4
  End If

  objReg.GetStringValue HKLM, strKey & strSubkey, _
  strEntry2, strValue2
  If strValue2 <> "" Then
    'WScript.Echo "Install Date: " & strValue2
  End If
  
Next

End Sub


'==================================================================
'==================================================================
'==================================================================




