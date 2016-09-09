If WScript.Arugments.Count <> 1 Then
	WScript.Echo "Usage: Convert2XLSX <Path_to_XLS_Files>"
	WScript.Quit(1)
End If

wd=WScript.Arugments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExist(wd) Then
	WScript.Echo "Error: Folder " & wd & " not exist or you don't have permission to access it!"
	WScript.Quit(2)
End If

Set app = CreateObject("Excel.Application")
Set folder = fso.GetFolder(wd)
For Each f In folder.Files
    If Right(f.Name, 4) = ".xls" Then
        Set wbk = app.Workbooks.Open(f)
        If wbk.HasVBProject Then
            wbk.SaveAs f & "m", 52
        Else
            wbk.SaveAs f & "x", 51
        End If
        wbk.Close False
    End If
Next
app.Quit(0)

