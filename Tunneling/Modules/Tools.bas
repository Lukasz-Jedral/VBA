Attribute VB_Name = "Tools"
Function File_List_In_Directory()
Dim UserInput As String, Folder_Path As String
Dim oFD As FileDialog, oFile As Object, oFiles As Object, oFolder As Object
Dim intChoice As Integer, i As Integer
Dim vaArray() As Variant
Dim Msg

UserInput = InputBox("List only files with specific extension?" & vbNewLine & "Leave empty to list all files", "List files in directory")


Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
oFD.Title = "Please select a folder contaning" & CStr(UserInput) & "files"
intChoice = oFD.Show

If intChoice <> 0 Then
    'get the folder path selected by the user
    Folder_Path = oFD.SelectedItems(1)
Else 'if no file will be pointed message will pop up
    Call MsgBox("Come on! Show me where Tun files are located.", _
    vbOKOnly + vbExclamation, "Don't fool around!")
    Exit Function
End If

Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(Folder_Path)
    Set oFiles = oFolder.Files


If oFiles.Count = 0 Then
    MsgBox ("Choosen folder is empty")
    Exit Function
End If

If Not UserInput = vbNullString Then
    i = 1
    For Each oFile In oFiles
        If LCase(Right(oFile.Name, 4)) = CStr(UserInput) Then
        ReDim Preserve vaArray(i)
        vaArray(i) = CStr(oFile.Path)
        i = i + 1
        End If
    Next
    
Else

ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = CStr(oFile.Path)
        i = i + 1
    Next
    
End If

File_List_In_Directory = vaArray


End Function

Function SelectFolder() As String

Dim diaFolder As FileDialog
'Open the file dialog
On Error GoTo ErrorHandler
Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
diaFolder.AllowMultiSelect = False
diaFolder.Title = "Select a folder then hit OK"
diaFolder.Show
SelectFolder = diaFolder.SelectedItems(1)
Set diaFolder = Nothing

Exit Function

ErrorHandler:
Msg = "No folder selected, you must select a folder for program to run"
Style = vbError
Title = "Need to Select Folder"
Response = MsgBox(Msg, Style, Title)

End Function

Function ExtractTunCoordinates(TunFile() As String) As String()
Dim Points() As String, Splitted() As String
Dim i As Integer, j As Integer, counter As Integer

ReDim Points(5, 0)

For i = 3 To UBound(TunFile)
    counter = 1
    Splitted = Split(TunFile(i), " ")
    
    For j = 0 To UBound(Splitted)
    
        If Splitted(j) <> "" Then
            ReDim Preserve Points(5, i - 2)
            Points(counter, i - 2) = Splitted(j)
            counter = counter + 1
        End If
    Next j
Next i

ExtractTunCoordinates = Points

End Function
