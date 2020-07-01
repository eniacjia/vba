'!!!!! USE ON YOUR OWN RISK !!!!!
'======= File Operation ==============================
' CREATOR: DJIA | eniacjia@gmail.com
' BUILD VERSION: [2020-07-01]
' FUNCTION: FSO File Operations
' PARAMETER LIST:
'   MyOperation: "COPY", "MOVE", "RENAME", "DELETE", "LIST"
'   SourceFile: Source Path
'   DestFile:  Dest Path
' UPDATE LOG ------------------------------------------ 
'  [2020-07-01] DJ: Update the descriptions
'======================================================
Function FileOperation(MyOperation, SourceFile, DestFile)
    Dim MyFso
    On Error Resume Next
    Set MyFso = CreateObject("Scripting.FileSystemObject")
    Select Case UCase(MyOperation)
        Case "COPY"
            MyFso.CopyFile SourceFile, DestFile, True '--- True means OverwriteExisting
        Case "MOVE"
            MyFso.CopyFile SourceFile, DestFile, True '--- True means OverwriteExisting
            MyFso.DeleteFile SourceFile
        Case "RENAME"
            MyFso.MoveFile SourceFile, DestFile
        Case "DELETE"
            MyFso.DeleteFile SourceFile
        Case "LIST"
            Dim MyFile, MyFolder, MyList
            Set MyFolder = MyFso.GetFolder(SourceFile)
            MyList = ""
            For Each MyFile In MyFolder.Files
                If InStr(UCase(MyFile.Name), UCase(DestFile)) > 0 Then
                    MyList = MyList & "  > " & MyFile.Name & "  (" & MyFile.Size & " | " & MyFile.DateCreated & ")" & vbCrLf
                End If
            Next
            FileOperation = MyList
        Case Else
            FileOperation = "UNKNOWN OPERATOIN: " & MyOperation
    End Select
    '-- If any error --------
    If Err.Number <> 0 Then
        FileOperation = Err.Description
    End If
    Set MyFso = Nothing
End Function
