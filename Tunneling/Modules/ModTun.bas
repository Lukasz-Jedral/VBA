Attribute VB_Name = "ModTun"
Sub Delete_Points_From_TUN()
Dim FileList As Variant
Dim FilePath As String, PointsToBeDeleted As String, LineFromFile As String, FileContent As String
Dim PointsSplitted() As String, RawGeoFile() As String
Dim counter As Integer, i As Integer, j As Integer, TextFile As Integer, PointNo As Integer
Dim FoundInFile As Boolean

FileList = Tools.File_List_In_Directory

PointsToBeDeleted = InputBox("Please type in numers of point to be deleted separated with comma")
PointsSplitted = Split(PointsToBeDeleted, ",")

FileContent = ""

For counter = 1 To UBound(FileList)
    FilePath = CStr(FileList(counter))
    i = 0
    ReDim RawGeoFile(0)
    
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open FilePath For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    
    Do Until EOF(TextFile) 'loop until end of file

        Line Input #TextFile, LineFromFile 'reads line from file
        
        FoundInFile = False
        
        For j = 0 To UBound(PointsSplitted) 'Splits UserInput and compares if givem point is in file
            If Left(LineFromFile, 2) = PointsSplitted(j) Then
                FoundInFile = True
            End If
        Next j
        
        If FoundInFile = False Then
            i = i + 1
            ReDim Preserve RawGeoFile(i)
            RawGeoFile(i) = LineFromFile
        End If

    Loop
    Close TextFile
        PointNo = 1
        For j = 3 To UBound(RawGeoFile)
           RawGeoFile(j) = Format(PointNo, "00") & Right(RawGeoFile(j), Len(RawGeoFile(j)) - 2)
           PointNo = PointNo + 1
        Next j
    'Kill (FilePath)
    
    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open FilePath For Output As TextFile
    For j = 1 To UBound(RawGeoFile)
        Print #TextFile, RawGeoFile(j)
    Next j
    Close TextFile
    
Next counter


End Sub

Sub ReNumberTUN()
Dim FileList() As Variant, RawGeoFile() As String
Dim FilePath As String, LineFromFile As String
Dim counter As Integer, TextFile As Integer, i As Integer, PointNo As Integer

FileList = Tools.File_List_In_Directory

    
For counter = 1 To UBound(FileList)
    FilePath = CStr(FileList(counter))
    i = 0
    ReDim RawGeoFile(0)
    
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open FilePath For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    Do Until EOF(TextFile) 'loop until end of file

        Line Input #TextFile, LineFromFile 'reads line from file
        i = i + 1
        ReDim Preserve RawGeoFile(i)
        RawGeoFile(i) = LineFromFile
    Loop
    Close TextFile
    
    PointNo = 1
    For j = 3 To UBound(RawGeoFile)
        RawGeoFile(j) = Format(PointNo, "00") & Right(RawGeoFile(j), Len(RawGeoFile(j)) - 2)
        PointNo = PointNo + 1
    Next j
    
    Open FilePath For Output As TextFile
    For j = 1 To UBound(RawGeoFile)
        Print #TextFile, RawGeoFile(j)
    Next j
    Close TextFile
    
Next counter

End Sub

Sub CombineTun()


Dim FileList1 As Variant, FileList2 As Variant
Dim FilePath1 As String, FilePath2 As String, NewFilePath As String, FilePath As String, LineFromFile As String
Dim counter As Integer, TextFile As Integer, i As Integer, PointNo As Integer, Response As Integer, PointNoLeft As Integer, PointNoRight As Integer
Dim TunFileLeft() As String, TunFileRight() As String, TunCombined() As String, FilePathSplitted() As String, FileName() As String

PointNoLeft = InputBox("Last point number on left side?")
FileList1 = Tools.File_List_In_Directory
PointNoRight = InputBox("First point on right side?")
FileList2 = Tools.File_List_In_Directory

If UBound(FileList1) <> UBound(FileList2) Then
    Response = MsgBox("Number of files to be connected needs to be the same in both indicated folders" & vbNewLine & "Macro will be closed now", 0)
    Exit Sub
End If
    
NewFilePath = Tools.SelectFolder

Debug.Print "Numbers of files in first folder: " & UBound(FileList1)
Debug.Print "Numbers of files in second folder: " & UBound(FileList2)

'For i = 1 To UBound(FileList1)
    'Debug.Print FileList1(i)
    'Debug.Print FileList2(i)
'Next i
    
For counter = 1 To UBound(FileList1)
    FilePath1 = CStr(FileList1(counter))
    FilePath2 = CStr(FileList2(counter))
    i = 0
    ReDim TunFileLeft(0)
    ReDim TunFileRight(0)
    '-------------extracting data from first file------------------------
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open FilePath1 For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    Do Until EOF(TextFile) 'loop until end of file

        Line Input #TextFile, LineFromFile 'reads line from file
        i = i + 1
        ReDim Preserve TunFileLeft(i)
        TunFileLeft(i) = LineFromFile
    Loop
    Close TextFile
    '-------------extracting data from second file------------------------
    i = 0
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open FilePath2 For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    Do Until EOF(TextFile) 'loop until end of file

        Line Input #TextFile, LineFromFile 'reads line from file
        i = i + 1
        ReDim Preserve TunFileRight(i)
        TunFileRight(i) = LineFromFile
    Loop
    Close TextFile
    '-------------combining files------------------------
    ReDim Preserve TunCombined(PointNoLeft + 2)
    TunCombined(1) = TunFileLeft(1)
    TunCombined(2) = TunFileLeft(2)
    PointNo = 1
    
    For j = 3 To PointNoLeft + 2
        TunCombined(j) = TunFileLeft(j)
        TunCombined(j) = Format(PointNo, "00") & Right(TunCombined(j), Len(TunCombined(j)) - 2)
        PointNo = PointNo + 1
    Next j
    
    ReDim Preserve TunCombined(PointNoLeft + 2 + (UBound(TunFileRight) - PointNoRight - 1))
    i = PointNoRight + 2
    
    For j = PointNoLeft + 3 To UBound(TunCombined)
        TunCombined(j) = TunFileRight(i)
        TunCombined(j) = Format(PointNo, "00") & Right(TunCombined(j), Len(TunCombined(j)) - 2)
        i = i + 1
        PointNo = PointNo + 1
    Next j
    
    'For j = 1 To UBound(TunCombined)
        'Debug.Print TunCombined(j)
    'Next j
    '-------------saving combinined file------------------------
    FilePathSplitted = Split(FilePath1, "\")
    FileName = Split(FilePathSplitted(UBound(FilePathSplitted)), ".")
    FilePath = NewFilePath & "\" & FileName(0) & "_combined." & FileName(1)
    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open FilePath For Output As TextFile

    'Write lines of text
    For j = 1 To UBound(TunCombined)
        Print #TextFile, TunCombined(j)
    Next j
    
  
    'Save & Close Text File
    Close TextFile
    
Next counter

MsgBox ("Number of crated files: " & counter - 1)
End Sub

Sub Renumber_Clockwise()
'makro will extract first and last point offset value. If first point offset is greater than last one it means that file needs to be rearranged
Dim FileList As Variant
Dim counter As Integer, TextFile As Integer, i As Integer, j As Integer, PointNo As Integer
Dim FilePath As String, LineFromFile As String
Dim TunFile() As String, FirstPoint() As String, LastPoint() As String, Splitted() As String, ReorganizedTun() As String

FileList = Tools.File_List_In_Directory

For counter = 1 To UBound(FileList)
    FilePath = CStr(FileList(counter))
    ReDim TunFile(0)
    i = 0
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open FilePath For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    
    Do Until EOF(TextFile) 'loop until end of file
        Line Input #TextFile, LineFromFile 'reads line from file
        i = i + 1
        ReDim Preserve TunFile(i)
        TunFile(i) = LineFromFile
    Loop
    Close TextFile
        'now all data from tun file is extracted and stored in TunFile variable
        'now we can to look at first point offset value and compare it to last point offset value
    
    
    Splitted = Split(TunFile(3), " ") 'extracting coords for first point
    i = 1
    For j = 0 To UBound(Splitted)
        If Splitted(j) <> "" Then
            ReDim Preserve FirstPoint(i)
            FirstPoint(i) = Splitted(j)
            i = i + 1
        End If
    Next j
    
    Splitted = Split(TunFile(UBound(TunFile)), " ") 'extracting coords for last point
    i = 1
    For j = 0 To UBound(Splitted)
        If Splitted(j) <> "" Then
            ReDim Preserve LastPoint(i)
            LastPoint(i) = Splitted(j)
            i = i + 1
        End If
    Next j
    
    If FirstPoint(3) > LastPoint(3) Then 'if condition is meet points will be reorganized - last will be first, first will be last
        
        ReDim ReorganizedTun(UBound(TunFile))
        ReorganizedTun(1) = TunFile(1)
        ReorganizedTun(2) = TunFile(2)
        
        i = 3
        
        For j = UBound(TunFile) To 3 Step -1 ' reversing the order of points
            ReorganizedTun(i) = TunFile(j)
            'Debug.Print ReorganizedTun(i)
            i = i + 1
        Next j
        
        PointNo = 1
        For j = 3 To UBound(ReorganizedTun)
            ReorganizedTun(j) = Format(PointNo, "00") & Right(ReorganizedTun(j), Len(ReorganizedTun(j)) - 2) 'renumbering points so it always starts from 01
            PointNo = PointNo + 1
        Next j
        
        
        Open FilePath For Output As TextFile 'overrighting reorganied file
            For j = 1 To UBound(ReorganizedTun)
                Print #TextFile, ReorganizedTun(j)
            Next j
        Close TextFile
        
    End If
    
    
Next counter

End Sub

Sub RemoveIdenticalProfiles()
Dim FileList As Variant
Dim counter As Integer, TextFile As Integer, i As Integer, j As Integer, Compare As Integer, ArrSize As Integer
Dim FilePath As String, NextFileFilePath As String, LineFromFile As String
Dim TunFile() As String, NextTunFile() As String, FilesToSave() As String, TunFilePoints() As String, NextTunFilePoints() As String


ReDim FilesToSave(0)
FileList = Tools.File_List_In_Directory

Compare = 1 'position on the list of file that is currently used as file to which we are comparing other files. It need to change when we find different files

For counter = 2 To UBound(FileList)
    FilePath = CStr(FileList(Compare))
    NextFileFilePath = CStr(FileList(counter))
    ReDim TunFile(0)
    ReDim NextTunFile(0)
    'Debug.Print NextFileFilePath
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open FilePath For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    
    i = 0
    Do Until EOF(TextFile) 'loop until end of file
        Line Input #TextFile, LineFromFile 'reads line from file
        i = i + 1
        ReDim Preserve TunFile(i)
        TunFile(i) = LineFromFile
    Loop
    Close TextFile
    
    TextFile = FreeFile 'Determine the next file number available for use by the FileOpen function
    
    Open NextFileFilePath For Input As TextFile 'opens file as file no. 1 in "to be read" mode
    
    i = 0
    Do Until EOF(TextFile) 'loop until end of file
        Line Input #TextFile, LineFromFile 'reads line from file
        i = i + 1
        ReDim Preserve NextTunFile(i)
        NextTunFile(i) = LineFromFile
    Loop
    Close TextFile
    
    
    'now data from two files were extracted. Now I will compare to files line by line. If they are different we will save both of them
    If UBound(TunFile) <> UBound(NextTunFile) Then 'if the size of both files are differnt it means they are different
        ArrSize = UBound(FilesToSave) 'chek how big FilesToSave array already is
        
        If counter - Compare = 1 Then
            ReDim Preserve FilesToSave(ArrSize + 1) 'making 1 new spots for saving new two file paths in FilesToSave array
            FilesToSave(ArrSize + 1) = FileList(Compare) & " " & CStr(UBound(TunFile) - 2)  'this step is to avoid double entries on list if next and previous files are different
            Compare = counter
            GoTo NextIteration
        Else
            ReDim Preserve FilesToSave(ArrSize + 2) 'making 2 new spots for saving new two file paths in FilesToSave array
            FilesToSave(ArrSize + 1) = FileList(Compare) & " " & CStr(UBound(TunFile) - 2)  'this step ensures that if we have series of indentical files first and last of them on list is saved
            FilesToSave(ArrSize + 2) = FileList(counter - 1) & " " & CStr(UBound(NextTunFile) - 2)
            Compare = counter
            GoTo NextIteration
        End If
    
    Else 'if both files have same size we will compare them line by line
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<     Extracting Points from Tun Files     >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
        TunFilePoints = Tools.ExtractTunCoordinates(TunFile)
        NextTunFilePoints = Tools.ExtractTunCoordinates(NextTunFile)
        
        For i = 1 To UBound(TunFilePoints, 2)
            If Abs(TunFilePoints(2, i) - NextTunFilePoints(2, i)) > 0.001 And Abs(TunFilePoints(3, i) - NextTunFilePoints(3, i)) > 0.001 Then 'if compared values are biiger than 1mm then it is assumed that files are not the same and so they need to be saved
                ArrSize = UBound(FilesToSave) 'chek how big FilesToSave array already is
                
                If counter - Compare = 1 Then
                    ReDim Preserve FilesToSave(ArrSize + 1) 'making 1 new spots for saving new two file paths in FilesToSave array
                    FilesToSave(ArrSize + 1) = FileList(Compare) & " " & CStr(UBound(TunFile) - 2) 'this step is to avoid double entries on list iv next and previous files are different
                    Compare = counter
                    GoTo NextIteration
                Else
                    ReDim Preserve FilesToSave(ArrSize + 2) 'making 2 new spots for saving new two file paths in FilesToSave array
                    FilesToSave(ArrSize + 1) = FileList(Compare) & " " & CStr(UBound(TunFile) - 2)  'this step ensures that if we have series of indentical files first and last of them on list is saved
                    FilesToSave(ArrSize + 2) = FileList(counter - 1) & " " & CStr(UBound(NextTunFile) - 2)
                    Compare = counter
                    GoTo NextIteration
                End If
            End If
        Next i
    
    End If
    
NextIteration:
Next counter

FilePath = Left(FileList(1), (InStrRev(FileList(1), "\"))) & "FilesToBeSaved.txt"
TextFile = FreeFile
Open FilePath For Output As TextFile

For i = 1 To UBound(FilesToSave)
    Print #TextFile, FilesToSave(i)
Next i

Close TextFile
End Sub
