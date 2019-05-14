
'====================================================================================================
'user variables : HKEY_CURRENT_USER\Environment.
'system variables : HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment.
'====================================================================================================

'Call ReadEnvVar("System", "PATH")
'Call ListAllInstalledApps()

Call GetJavaWmiCPanelReg ()
Call GetJavaWmiEnvVars



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


'--------------------------------------------------
' **** LIST INSTALLED JDK/JRE From REGISTRY **** '
'--------------------------------------------------
Sub GetJavaWmiCPanelReg()

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."
strCallType = "CPanel"

strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry2 = "InstallDate"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"

Dim strAppName, strAppVersion, strAppDate
Dim arrJDK, arrJRE
Dim arrJDKTemp(), arrJRETemp()
strJDKFound = 0
strJREFound = 0

Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
 
objReg.EnumKey HKLM, strKey, arrSubkeys
'WScript.Echo "Installed Applications" & vbCrLf

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
    
    If (InStr(strAppName, "Java") > 0) Then
        Select Case IsJdkJreString(strAppName, strCallType)
            Case "JDK"
                arrJDK = ArrayFill(arrJDKTemp, strAppName, strAppVersion, strAppDate, strJDKFound)
                strJDKFound = strJDKFound + 1
            Case "JRE"
                arrJRE = ArrayFill(arrJRETemp, strAppName, strAppVersion, strAppDate, strJREFound)
                strJREFound = strJREFound + 1
        End Select
    End If

  End If
  
Next

  If strJDKFound = 0 Then
    WScript.Echo "No JDK Found Installed"
  Else
    Call PublishJDKJRE(arrJDK, 2)
  End If
  
  If strJREFound = 0 Then
    WScript.Echo "No JRE Found Installed"
  Else
    Call PublishJDKJRE(arrJRE, 2)
  End If
 
    Set objReg = Nothing

End Sub

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

'-----------------------------------------
' **** FILL JDK/JRE ARRAY **** '
'-----------------------------------------
Function ArrayFill(ByRef arrJDKJRE, strAppName, strAppVersion, strAppDate, strJDKJREFound)

ReDim Preserve arrJDKJRE(2, strJDKJREFound)
arrJDKJRE(0, strJDKJREFound) = strAppName
arrJDKJRE(1, strJDKJREFound) = strAppVersion
arrJDKJRE(2, strJDKJREFound) = strAppDate
ArrayFill = arrJDKJRE

End Function

'-----------------------------------------
' **** PUBLISH OUTPUT **** '
'-----------------------------------------
Function PublishJDKJRE(arrJDKJRE, strDimension)

Dim counter, i

Select Case strDimension
    Case 1
        For i = 0 To UBound(arrJDKJRE)
            WScript.Echo arrJDKJRE(i)
        Next
         WScript.Echo vbCrLf
    Case 2
        For counter = 0 To UBound(arrJDKJRE, 2)
            WScript.Echo arrJDKJRE(0, counter) 
            WScript.Echo arrJDKJRE(1, counter) 
            WScript.Echo arrJDKJRE(2, counter) & vbCrLf
        Next
End Select

End Function



'--------------------------------------------------
' **** LIST JDK/JRE ENV VARAIABLES **** '
'--------------------------------------------------
Sub GetJavaWmiEnvVars()
'Using WMI retrieves both USER and SYSTEM variable together, you cannot pick and choose
strCallType = "javapath"

Dim arrJDKSys, arrJRESys
strCntJDKSys = 0
strCntJRESys = 0

Dim arrJDKUsr, arrJREUsr
strCntJDKUsr = 0
strCntJREUsr = 0

Dim arrPathUsr, arrPathSys
strCntPathUsr = 0
strCntPathSys = 0

strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")

For Each objItem In colItems

    If (StrComp(objItem.Name, "JAVA_HOME") = 0) Then
        Select Case (objItem.SystemVariable)
            Case "True"
                arrJDKSys = Array(objItem.Name, objItem.VariableValue)
                strCntJDKSys = strCntJDKSys + 1
            Case "False"
                arrJDKUsr = Array(objItem.Name, objItem.VariableValue)
                strCntJDKUsr = strCntJDKUsr + 1
        End Select
    ElseIf (StrComp(objItem.Name, "JRE_HOME") = 0) Then
        Select Case (objItem.SystemVariable)
            Case "True"
                arrJRESys = Array(objItem.Name, objItem.VariableValue)
                strCntJRESys = strCntJDKSys + 1
            Case "False"
                arrJREUsr = Array(objItem.Name, objItem.VariableValue)
                strCntJREUsr = strCntJDKUsr + 1
        End Select
    End If
    
    If (StrComp(objItem.Name, "Path") = 0) Then
        If (IsJdkJreString(objItem.VariableValue, strCallType) <> "False") Then
            Select Case (objItem.SystemVariable)
                Case "True"
                    arrPathSys = Array(objItem.Name, ExtractPathValue(objItem.VariableValue, strCallType))
                    strCntPathSys = strCntPathSys + 1
                Case "False"
                    arrPathUsr = Array(objItem.Name, ExtractPathValue(objItem.VariableValue, strCallType))
                    strCntPathUsr = strCntPathUsr + 1
            End Select
        Else
            strCntPathSys = 0
            strCntPathUsr = 0
        End If
    End If
Next

arrStrOut = Array(arrJDKSys, arrJRESys, arrJDKUsr, arrJREUsr, arrPathUsr, arrPathSys)
For Each arrFound In arrStrOut
    If IsArray(arrFound) Then
        If (arrFound(0) <> "") Then
            Call PublishJDKJRE(arrFound, "1")
        End If
    End If
Next

Set colItems = Nothing
Set objWMIService = Nothing

End Sub



'--------------------------------------------------
' **** EXTRACT VALUE FROM SYS VARIABLES **** '
'--------------------------------------------------
Function ExtractPathValue(strFullPathValue, strValueType)

strJavaPath = ""
strStatPosRight = InStr(strFullPathValue, strValueType) + (Len(strValueType) - 1)
strNew = Mid(strFullPathValue, 1, strStatPosRight)
cntExtract = Len(strNew)

For i = 1 To Len(strNew)
    If (Mid(strNew, cntExtract, 1) = ";") Then
        Exit For
    Else
        strJavaPath = strJavaPath & Mid(strNew, cntExtract, 1)
        cntExtract = cntExtract - 1
    End If
Next

ExtractPathValue = StrReverse(strJavaPath)

End Function




Sub GetJavaWmiJSoftReg()

End Sub



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




