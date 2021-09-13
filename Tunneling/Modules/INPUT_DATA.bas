Attribute VB_Name = "INPUT_DATA"
Option Explicit
Option Base 1

Function GetTheFilePath()

'defining variables
Dim File_Path As String
Dim UserID As String
Dim intChoice As Integer

'only allow the user to select one file
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
'change the display name of the open file dialog
Application.FileDialog(msoFileDialogOpen).Title = "Select Coordinate File"
'Remove all other filters
Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
'Add a custom filter
Call Application.FileDialog(msoFileDialogOpen).Filters.Add("Coordinate Document", "*.geo")
'make the file dialog visible to the user
intChoice = Application.FileDialog(msoFileDialogOpen).Show
'determine what choice the user made
If intChoice <> 0 Then
    'get the file path selected by the user
    File_Path = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else 'if no file will be pointed message will pop up
    Call MsgBox("Come on! Show me where the file is.", _
    vbOKOnly + vbExclamation, "Don't fool around!")
    Exit Function
End If

GetTheFilePath = File_Path

End Function
Function GetTheFilePathTun()

'defining variables
Dim File_Path As String
Dim UserID As String
Dim intChoice As Integer

'only allow the user to select one file
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
'change the display name of the open file dialog
Application.FileDialog(msoFileDialogOpen).Title = "Select Coordinate File"
'Remove all other filters
Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
'Add a custom filter
Call Application.FileDialog(msoFileDialogOpen).Filters.Add("Theoretic Tunnel", "*.tun")
'make the file dialog visible to the user
intChoice = Application.FileDialog(msoFileDialogOpen).Show
'determine what choice the user made
If intChoice <> 0 Then
    'get the file path selected by the user
    File_Path = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else 'if no file will be pointed message will pop up
    Call MsgBox("Come on! Show me where the file is.", _
    vbOKOnly + vbExclamation, "Don't fool around!")
    Exit Function
End If

GetTheFilePathTun = File_Path

End Function

Sub CoordinatesFromGeoFile(ByRef Points() As String, Lines() As String, File_Path As String)
'---------------------------------------------------
'Dim File_Path As String
'File_Path = "Z:\LOC\SEVAY10\03_Execution\SE_E4_Johannelund\02 Execution Docs\2.07 Surveying\12 KOLLEGOR\Lukasz\FSE403\03_Inmätning\TEST.geo"
'-----------------------------------------------------

Dim LineFromFile As String, LineItems As String
Dim RawGeoFile() As String, TempSplited() As String, RawLines() As String, RawPoints() As String ', Points() As String, Lines() As String
Dim i As Integer, j As Integer, k As Integer, n As Integer, StartPos As Integer, EndPos As Integer, LineNo As Integer

If File_Path = "" Then
Call MsgBox("Incorrect file path. Please load correct file and try again", vbCritical + vbOKCancel, "That will not work!")
Exit Sub
Else
Open File_Path For Input As #1 'opens file as file no. 1 in "to be read" mode
End If

Do Until EOF(1) 'loop until end of file

Line Input #1, LineFromFile 'reads line from file no. 1

i = i + 1
ReDim Preserve RawGeoFile(i)
RawGeoFile(i) = LineFromFile
RawGeoFile(i) = Replace(RawGeoFile(i), vbTab, "")

Loop
Close #1

'vvvvvvvvv coping from the RawGeoFile part containing only lines vvvvvvvvvvv

On Error Resume Next 'errors from now on will be ignored and next line of the code will be complied

For i = 1 To UBound(RawGeoFile) - 1
    If RawGeoFile(i) = "LineList " And RawGeoFile(i + 1) = "begin" Then 'error can accure in here if there is no begin after LineList w can get i value bigger than array dimension
        For j = 1 To (UBound(RawGeoFile) - i)
            ReDim Preserve RawLines(j)
            RawLines(j) = RawGeoFile(i + k)
            k = k + 1
        Next j
        Exit For
    End If
Next i

On Error GoTo 0 'normal error handling is now restored

'For k = 1 To UBound(RawLines)
'Debug.Print RawLines(k)
'Next k

ReDim Preserve RawGeoFile(UBound(RawGeoFile) - j)

'vvvvvvvvvvvv coping from the RawGeoFile part containing only points vvvvvvvvvvvvvvvvv

On Error Resume Next 'errors from now on will be ignored and next line of the code will be complied

For i = 1 To UBound(RawGeoFile) - 1
    If RawGeoFile(i) = "PointList " And RawGeoFile(i + 1) = "begin" Then
        j = 1
        Do Until RawGeoFile(i + j + 1) = "end"
            ReDim Preserve RawPoints(j)
            RawPoints(j) = RawGeoFile(i + j + 1)
            j = j + 1
        Loop
        Exit For
    End If
Next i

On Error GoTo 0 'normal error handling is now restored

'vvvvvvvvvvvvvvvvvvvvvvvv Extracting Points into array vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv

If UBound(RawPoints) <> 0 Then
j = 1
    For i = 1 To UBound(RawPoints)
        TempSplited = Split(RawPoints(i), ",", , vbBinaryCompare)
        ReDim Preserve Points(4, j)
            For k = 1 To 4
            Points(k, j) = TempSplited(k - 1)
            Next k
        StartPos = InStr(1, Points(1, j), Chr(34), vbBinaryCompare)
        EndPos = InStrRev(Points(1, j), Chr(34), , vbBinaryCompare)
        Points(1, j) = Mid(Points(1, j), StartPos + 1, EndPos - StartPos - 1)
        j = j + 1
    Next i
End If

'vvvvvvvvvvvvvvvvvvvvvvvv Extracting Lines into array vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv

If UBound(RawLines) <> 0 Then

    LineNo = 1
    j = 1
    
    For i = 1 To UBound(RawLines) - 1
        If RawLines(i) = "PointList " And RawLines(i + 1) = "begin" Then 'error can accure in here if there is no begin after LineList w can get i value bigger than array dimension
            
            k = i + 2
            
            Do While RawLines(k) <> "end"
                TempSplited = Split(RawLines(k), ",", , vbBinaryCompare)
                ReDim Preserve Lines(5, j)
                    For n = 1 To 5
                        If n = 5 Then
                            Lines(n, j) = "Line" & LineNo
                        Else
                            Lines(n, j) = TempSplited(n - 1)
                        End If
                    Next n
                StartPos = InStr(1, Lines(1, j), Chr(34), vbBinaryCompare)
                EndPos = InStrRev(Lines(1, j), Chr(34), , vbBinaryCompare)
                Lines(1, j) = Mid(Lines(1, j), StartPos + 1, EndPos - StartPos - 1)
                j = j + 1
                k = k + 1
            Loop
            
            LineNo = LineNo + 1
            
        End If
    Next i
End If

Call MsgBox("Imported:" & vbNewLine & vbNewLine & "Points:  " & UBound(Points, 2) & vbNewLine & "Lines:      " & LineNo - 1, vbOKOnly)

End Sub

Function Import_Lines(File_Path As String) As String()

Dim LineFromFile As String, LineStatus As String
Dim RawGeoFile() As String, RawLines() As String, TempSplited() As String, Lines() As String
Dim i As Integer, n As Integer, LineListPos As Integer, LineCount As Integer, StartPos As Integer, EndPos As Integer, ArrSize As Integer

'File_Path = "Z:\LOC\SEVAY10\03_Execution\SE_E4_Johannelund\02 Execution Docs\2.07 Surveying\12 KOLLEGOR\Lukasz\FSE403\03_Inmätning\TEST.geo"

If File_Path = "" Then
Call MsgBox("Incorrect file path. Please load correct file and try again", vbCritical + vbOKCancel, "That will not work!")
Exit Function
Else
Open File_Path For Input As #1 'opens file as file no. 1 in "to be read" mode
End If

Do Until EOF(1) 'loop until end of file

Line Input #1, LineFromFile 'reads line from file no. 1

i = i + 1
ReDim Preserve RawGeoFile(i)
RawGeoFile(i) = LineFromFile
RawGeoFile(i) = Replace(RawGeoFile(i), vbTab, "")

Loop
Close #1

LineListPos = INPUT_DATA.String_Position(RawGeoFile(), "LineList ")

RawLines() = INPUT_DATA.Copy_Part_Of_Array(RawGeoFile(), LineListPos + 2)

For i = 1 To UBound(RawLines)
    TempSplited = Split(RawLines(i), ",", , vbBinaryCompare)
    If UBound(TempSplited) <> 0 Then
        If InStr(1, TempSplited(0), "Line", vbBinaryCompare) Then
            If TempSplited(1) = "1" Then
                LineCount = LineCount + 1
                LineStatus = "Closed"
            Else
                LineCount = LineCount + 1
                LineStatus = "Open"
            End If
        End If

        If InStr(1, TempSplited(0), "Point ", vbBinaryCompare) Then
            ArrSize = ArrSize + 1
            ReDim Preserve Lines(6, ArrSize)
            
                    For n = 1 To 4
                            Lines(n, ArrSize) = TempSplited(n - 1)
                    Next n
                    
                StartPos = InStr(1, Lines(1, ArrSize), Chr(34), vbBinaryCompare)
                EndPos = InStrRev(Lines(1, ArrSize), Chr(34), , vbBinaryCompare)
                Lines(1, ArrSize) = Mid(Lines(1, ArrSize), StartPos + 1, EndPos - StartPos - 1)
            Lines(5, ArrSize) = "Line " & LineCount
            Lines(6, ArrSize) = LineStatus
        End If
    End If
Next i

Import_Lines = Lines

End Function

Function String_Position(Object_to_be_searched() As String, String_being_sought As String) As Integer

Dim i As Integer

For i = 1 To UBound(Object_to_be_searched) ' finding posistion of String_being_sought in Object_to_be_searched
    If Object_to_be_searched(i) = String_being_sought Then
        String_Position = i
        Exit For
    End If
Next i

End Function

Function Copy_Part_Of_Array(Source_Array() As String, Optional Starting As Integer, Optional Ending As Integer) As String()

Dim i As Integer, j As Integer
Dim New_Array() As String

If Starting = 0 Then
    Starting = LBound(Source_Array)
End If

If Ending = 0 Then
    Ending = UBound(Source_Array)
End If
    
j = 1
    
For i = Starting To Ending
    ReDim Preserve New_Array(j)
    New_Array(j) = Source_Array(i)
    j = j + 1
Next i

Copy_Part_Of_Array = New_Array

End Function
Function Import_Points(File_Path As String) As String()

Dim i As Integer, n As Integer, LineListPos As Integer, PointListPos As Integer, ArrSize As Integer, StartPos As Integer, EndPos As Integer
Dim LineFromFile As String
Dim RawGeoFile() As String, RawPoints() As String, TempSplited() As String, Points() As String

'File_Path = "Z:\LOC\SEVAY10\03_Execution\SE_E4_Johannelund\02 Execution Docs\2.07 Surveying\12 KOLLEGOR\Lukasz\FSE403\03_Inmätning\TEST.geo"

If File_Path = "" Then
Call MsgBox("Incorrect file path. Please load correct file and try again", vbCritical + vbOKCancel, "That will not work!")
Exit Function
Else
Open File_Path For Input As #1 'opens file as file no. 1 in "to be read" mode
End If

Do Until EOF(1) 'loop until end of file

Line Input #1, LineFromFile 'reads line from file no. 1

i = i + 1
ReDim Preserve RawGeoFile(i)
RawGeoFile(i) = LineFromFile
RawGeoFile(i) = Replace(RawGeoFile(i), vbTab, "")

Loop
Close #1

LineListPos = INPUT_DATA.String_Position(RawGeoFile(), "LineList ")

PointListPos = INPUT_DATA.String_Position(RawGeoFile(), "PointList ")

RawPoints() = INPUT_DATA.Copy_Part_Of_Array(RawGeoFile(), PointListPos, LineListPos)

For i = 1 To UBound(RawPoints)
    TempSplited = Split(RawPoints(i), ",", , vbBinaryCompare)
        If InStr(1, TempSplited(0), "Point ", vbBinaryCompare) Then
            ArrSize = ArrSize + 1
            ReDim Preserve Points(4, ArrSize)
            
                    For n = 1 To 4
                            Points(n, ArrSize) = TempSplited(n - 1)
                    Next n
                    
                StartPos = InStr(1, Points(1, ArrSize), Chr(34), vbBinaryCompare)
                EndPos = InStrRev(Points(1, ArrSize), Chr(34), , vbBinaryCompare)
                Points(1, ArrSize) = Mid(Points(1, ArrSize), StartPos + 1, EndPos - StartPos - 1)
        End If
Next i

Import_Points = Points

End Function

Function Get_Folder_Content(InitialFolderName As String)

Dim oFile As Object
Dim oFSO As Object
Dim oFolder As Object
Dim oFiles As Object
Dim intChoice As Integer, i As Integer
Dim vaArray As Variant
Dim oFD As FileDialog
Dim Folder_Path As String


Set oFD = Application.FileDialog(msoFileDialogFolderPicker)
oFD.Title = "Please Select a Folder contaning .tun files"
oFD.InitialFileName = InitialFolderName
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

ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = CStr(oFile.Path)
        i = i + 1
    Next


Get_Folder_Content = vaArray


End Function

Function SaveFilePath()

'defining variables
Dim File_Path As String
Dim UserID As String
Dim intChoice As Integer
Dim oFD As FileDialog

Set oFD = Application.FileDialog(msoFileDialogSaveAs)

oFD.Title = "Create New TBS file"
oFD.InitialFileName = "Z:\LOC\SEVAY10\03_Execution\SE_E4_Johannelund\02 Execution Docs\2.07 Surveying\NewTBS" 'sets the initial folder and file name

intChoice = oFD.Show

If intChoice <> 0 Then
    'get the file path selected by the user
    File_Path = Application.FileDialog(msoFileDialogSaveAs).SelectedItems(1)
Else 'if no file will be pointed message will pop up
    Call MsgBox("Come on! Show me where to save file.", _
    vbOKOnly + vbExclamation, "Don't fool around!")
    Exit Function
End If

SaveFilePath = File_Path

End Function

Sub CreateAfile(File_Path As String)

Dim fso As FileSystemObject
Dim fileStream As TextStream

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.CreateTextFile(File_Path, True) 'this line creates file
    fileStream.Close
    
End Sub

Function GetTheFilePathExtended(WinTitle As String, FileDescription As String, FileExtension As String, InitialFolder As String)

'defining variables
Dim File_Path As String
Dim UserID As String
Dim intChoice As Integer

'only allow the user to select one file
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
'change the display name of the open file dialog
Application.FileDialog(msoFileDialogOpen).Title = WinTitle
'Remove all other filters
Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
'Add a custom filter
Call Application.FileDialog(msoFileDialogOpen).Filters.Add(FileDescription, FileExtension)

Application.FileDialog(msoFileDialogOpen).InitialFileName = InitialFolder

'make the file dialog visible to the user
intChoice = Application.FileDialog(msoFileDialogOpen).Show
'determine what choice the user made
If intChoice <> 0 Then
    'get the file path selected by the user
    File_Path = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Else 'if no file will be pointed message will pop up
    Call MsgBox("Come on! Show me where the file is.", _
    vbOKOnly + vbExclamation, "Don't fool around!")
    Exit Function
End If

GetTheFilePathExtended = File_Path

End Function

Function SelectMultipleTun()

'defining variables
Dim File_Path() As String
Dim UserID As String
Dim fd As FileDialog
Dim fileChosen As Integer
Dim i As Integer
Set fd = Application.FileDialog(msoFileDialogFilePicker)



'only allow the user to select one file
fd.AllowMultiSelect = True
'change the display name of the open file dialog
fd.Title = "Select Tun Files"
'Remove all other filters
Call fd.Filters.Clear
'Add a custom filter
Call fd.Filters.Add("Tun Profiles", "*.tun")
'make the file dialog visible to the user
fileChosen = fd.Show
'determine what choice the user made
If fileChosen = -1 Then
    'get the file path selected by the user
    For i = 1 To fd.SelectedItems.Count
            'Debug.Print (fd.SelectedItems(i))
            ReDim Preserve File_Path(i)
            File_Path(i) = fd.SelectedItems(i)
            'Debug.Print (fileName)
        Next i
Else 'if no file will be pointed message will pop up
    Call MsgBox("Come on! Show me where the file is.", _
    vbOKOnly + vbExclamation, "Don't fool around!")
    Exit Function
End If

SelectMultipleTun = File_Path

End Function

