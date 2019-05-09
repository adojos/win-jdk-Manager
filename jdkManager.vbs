
'Call ReadEnvVar("System", "PATH")
'Call ListAllInstalledApps()
Call ListInstalledJdkJre ()



'-----------------------------------------
' **** READING ENVIRONMENT VARIABLE **** '
'-----------------------------------------
Sub ReadEnvVar(envVarType, envVarName)

Dim arrStrVar, strVar
    Set objShell = CreateObject("WScript.Shell")
    Set objEnv = objShell.Environment(envVarType)
    'WScript.Echo objEnv(envVarName)
    arrStrVar = Split(objEnv(envVarName), ";")

    For Each strVar In arrStrVar
        WScript.Echo strVar
    Next

End Sub

'-----------------------------------------
' **** WRITING ENVIRONMENT VARIABLE **** '
'-----------------------------------------
Sub WriteEnvVar(envVarType, envVarName, envVarValue)

Set objShell = CreateObject("WScript.Shell")
Set objEnv = objShell.Environment(envVarType)
objEnv(envVarName) = envVarValue

End Sub


'-----------------------------------------
' **** LIST INSTALLED JDK/JRE **** '
'-----------------------------------------
Sub ListInstalledJdkJre()

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."

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
        Select Case IsJdkJreString(strAppName)
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
    PublishJDKJRE (arrJDK)
  End If
  
  If strJREFound = 0 Then
    WScript.Echo "No JRE Found Installed"
  Else
    PublishJDKJRE (arrJRE)
  End If
  

End Sub


'---------------------------------------------------
' **** EVALUATE IF JDK/JRE FOUND IN APPNAME **** '
'---------------------------------------------------
Function IsJdkJreString(strJdkJre)

strJDK = "Development Kit"

If (InStr(1, strJdkJre, strJDK, 1) <> 0) Then
    IsJdkJreString = "JDK"
ElseIf (IsNumeric(Mid(strJdkJre, 6, 1))) Then
    IsJdkJreString = "JRE"
Else
    IsJdkJreString = False
End If

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
Function PublishJDKJRE(arrJDKJRE)

For counter = 0 To UBound(arrJDKJRE, 2)
    WScript.Echo arrJDKJRE(0, counter)
    WScript.Echo arrJDKJRE(1, counter)
    WScript.Echo arrJDKJRE(2, counter)
Next

End Function


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


