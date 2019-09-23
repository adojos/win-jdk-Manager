
Set obj = CreateLogFile()
call WriteLogFile(obj, "Hello World")




Public Function WriteLogFile (objFSOHandle, strMsg)

'Dim strDirName
'Dim strFileName
'strDirName = "C:\EuroDataGen"
'strFileName = "C:\EuroGenLog" & "-" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2) & ".txt"

'Set ObjFSO = CreateObject("Scripting.FileSystemObject")
'Set ObjTextFile = ObjFSO.OpenTextFile(strDirName & strFileName, 8, True)
objFSOHandle.WriteLine (strMsg & vbCrLf)

'Set ObjTextFile = Nothing
'Set ObjFSO = Nothing

End Function

'###########################################################################
'This function sets the directory path of ReadMe.txt 

Public Function CreateLogFile ()

sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
strFileName = "vbsx-Validator" & "_" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2) & ".txt"

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(sCurrPath & strFileName, 8, True)

Set CreateLogFile = ObjTextFile 

Set ObjTextFile = Nothing
Set ObjFSO = Nothing

'The Other Methods -
'Dim sCurrPath
'sCurrPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
'Msgbox sCurrPath

End Function