Attribute VB_Name = "SEC_TO_TUN"
Option Explicit
Option Base 1

Sub SEC_TO_TUN()

Dim i As Integer, j As Integer, n As Integer, ArrSize As Integer, NoOfFiles As Integer
Dim spaceA As Long, spaceB As Long, spaceC As Long, spaceD As Long
Dim File_Path As String, NewFilePath As String, LineFromFile As String, TodaysDate As String, Prefix As String, a As String, b As String, c As String, D As String
Dim RawGeoFile() As String, TempSplited() As String, CrossSectionPoints() As String, TunFilePoints() As String, TunFile() As String
Dim mean As Double
Dim fso As FileSystemObject
Set fso = New FileSystemObject
Dim fileStream As TextStream
Const myDecSep As String = "."

TodaysDate = Format(Now, "YYYY-MM-DD")
NoOfFiles = 0

File_Path = INPUT_DATA.GetTheFilePath
Prefix = InputBox("Prefix for new TUN files")

If File_Path = "" Then
    Call MsgBox("No file loaded!", vbOKOnly + vbCritical)
    Exit Sub
End If

Open File_Path For Input As #1 'opens file as file no. 1 in "to be read" mode


Do Until EOF(1) 'loop until end of file

Line Input #1, LineFromFile 'reads line from file no. 1

i = i + 1
ReDim Preserve RawGeoFile(i)
RawGeoFile(i) = LineFromFile
RawGeoFile(i) = Replace(RawGeoFile(i), vbTab, "")

Loop
Close #1

'this part of loop is serching for expresion Line " when it occurs it will start to extract coordiates
'for this particular line and after its done continiue to the next
For i = 1 To UBound(RawGeoFile)

    If InStr(1, RawGeoFile(i), "Line ", vbTextCompare) Then 'after finding expresion we search for makro will extrat coordinate reorganize them and save them as .tun file
        j = i + 4
        ArrSize = 1
        
        Do While InStr(1, RawGeoFile(j), "end", vbTextCompare) = 0
            TempSplited = Split(RawGeoFile(j), ",", , vbBinaryCompare)
            ReDim Preserve CrossSectionPoints(3, ArrSize)
            
            For n = 1 To 3
                CrossSectionPoints(n, ArrSize) = TempSplited(n)
            Next n
            
            
            ArrSize = ArrSize + 1
            j = j + 1
        Loop
        mean = 0
        'now when we did extract sektion, offset and hight for each point we can prepare tun file
        'calculating mean value for sektion for naming purposes
        For n = 1 To UBound(CrossSectionPoints, 2)
            ReDim Preserve TunFilePoints(5, n)
            TunFilePoints(1, n) = Format(n, "0#")
            TunFilePoints(2, n) = CrossSectionPoints(3, n)
            TunFilePoints(3, n) = CrossSectionPoints(2, n)
            TunFilePoints(4, n) = CrossSectionPoints(3, n)
            TunFilePoints(5, n) = ","
            mean = mean + CDbl(CrossSectionPoints(1, n))
        Next n
            mean = mean / UBound(CrossSectionPoints, 2)
            ReDim Preserve TunFile(2)
            TunFile(1) = "XYZ-COORD-FILE  ,V1.00," & TodaysDate & "," & "                                        ,"
            TunFile(2) = "                                        ,                                 ,"
            ArrSize = UBound(TunFilePoints, 2) + 2
            
        For n = 3 To ArrSize
            ReDim Preserve TunFile(n)
            a = TunFilePoints(1, n - 2)
            b = Replace(Format(TunFilePoints(2, n - 2), "#0.0000"), Application.DecimalSeparator, myDecSep) 'little trick for displaing dot as decimal sign regardless regional settings
            c = Replace(Format(TunFilePoints(3, n - 2), "#0.0000"), Application.DecimalSeparator, myDecSep)
            D = Replace(Format(TunFilePoints(4, n - 2), "#0.0000"), Application.DecimalSeparator, myDecSep)
            
            spaceA = 26 - (Len(a) + Len(b)) 'calculating number of space character so file formatting is readable for GEO
            spaceB = 38 - (Len(a) + spaceA + Len(b) + Len(c))
            spaceC = 50 - (Len(a) + spaceA + Len(b) + spaceB + Len(c) + Len(D))
            spaceD = 74 - (Len(a) + spaceA + Len(b) + spaceB + spaceC + Len(c) + Len(D))
            
            TunFile(n) = a & Space(spaceA) & _
                        b & Space(spaceB) & _
                        c & Space(spaceC) & _
                        D & Space(spaceD) & _
                        TunFilePoints(5, n - 2)
            'Debug.Print TunFile(n)
        Next n
            
            
            NewFilePath = Left(File_Path, InStrRev(File_Path, "\")) & Prefix & "_" & Format(mean, "#0.000") & ".tun"
            Set fileStream = fso.CreateTextFile(NewFilePath)
        
        For n = 1 To UBound(TunFile)
            fileStream.WriteLine TunFile(n)
        Next n
            fileStream.Close
        NoOfFiles = NoOfFiles + 1
    End If
    
Next i

MsgBox ("All Done! Number of created TUN files = " & NoOfFiles)
End Sub
Sub Add_Radius_to_TUN()
Dim File_Path As String, LineFromFile As String, Folder_Path As String
Dim RawTunFile() As String, Splitted() As String, Coords() As String, TunFile() As String
Dim i As Integer, j As Integer, n As Integer, ArrSize As Integer, NoOfFiles As Integer, counter As Integer
Dim dist_a As Double, dist_c As Double, dist_b As Double, area As Double, radius As Double
Dim spaceA As Long, spaceB As Long, spaceC As Long, spaceD As Long, spaceE As Long
Dim TunFileList() As Variant
Dim fso As FileSystemObject
Set fso = New FileSystemObject
Dim fileStream As TextStream



Folder_Path = "Z:\LOC\SEVAY10\03_Execution\SE_E4_Johannelund\02 Execution Docs\2.07 Surveying\05 TUNNEL\402\02 FÖRBEREDELSE\07 Membran\504"

TunFileList = INPUT_DATA.Get_Folder_Content(Folder_Path) 'getting file paths for tun files

NoOfFiles = 0

For counter = 1 To UBound(TunFileList)
    i = 0
    File_Path = TunFileList(counter)
    Debug.Print File_Path
    
    If File_Path = "" Then
        Call MsgBox("No file loaded!", vbOKOnly + vbCritical)
        Exit Sub
    End If
    'Debug.Print File_Path
    Open File_Path For Input As #1 'opens file as file no. 1 in "to be read" mode


    Do Until EOF(1) 'loop until end of file

    Line Input #1, LineFromFile 'reads line from file no. 1

    i = i + 1
    ReDim Preserve RawTunFile(i)
    RawTunFile(i) = LineFromFile
    RawTunFile(i) = Replace(RawTunFile(i), vbTab, "")
    'Debug.Print RawTunFile(i)
    Loop
    Close #1

    'ReDim TunFile(2)
    'TunFile(1) = RawTunFile(1)
    'TunFile(2) = RawTunFile(2)

    ArrSize = 1
    ReDim Coords(4, ArrSize)
    'extracting coordinates from RawTunFile. File is splited with space character. This creates many empty valuse. This code extraxts point number, height and offset from file.
    For i = 3 To UBound(RawTunFile)
        Splitted = Split(RawTunFile(i), " ")
        ReDim Preserve Coords(4, ArrSize)
        n = 1
        For j = 0 To UBound(Splitted)
            If Splitted(j) <> "" Then
                Coords(n, i - 2) = Splitted(j)
                n = n + 1
                If n = 4 Then
                    Exit For
                End If
            End If
        Next j
        ArrSize = ArrSize + 1
        Debug.Print Coords(1, i - 2) & " " & Coords(2, i - 2) & " " & Coords(3, i - 2)
    Next i

    'calculating radius
    For i = 2 To UBound(Coords, 2) - 2
        dist_a = Sqr((CDbl(Coords(2, i + 1)) - CDbl(Coords(2, i))) ^ 2 + ((Coords(3, i + 1)) - CDbl(Coords(3, i))) ^ 2)
        dist_b = Sqr((CDbl(Coords(2, i + 2)) - CDbl(Coords(2, i + 1))) ^ 2 + ((Coords(3, i + 2)) - CDbl(Coords(3, i + 1))) ^ 2)
        dist_c = Sqr((CDbl(Coords(2, i)) - CDbl(Coords(2, i + 2))) ^ 2 + ((Coords(3, i)) - CDbl(Coords(3, i + 2))) ^ 2)
    
        area = (CDbl(Coords(2, i)) * (CDbl(Coords(3, i + 1)) - CDbl(Coords(3, i + 2))) + CDbl(Coords(2, i + 1)) * (CDbl(Coords(3, i + 2)) - CDbl(Coords(3, i))) + CDbl(Coords(2, i + 2)) * (CDbl(Coords(3, i)) - CDbl(Coords(3, i + 1)))) / 2
    
        radius = (dist_a * dist_b * dist_c) / (4 * area)
        Coords(4, i) = "R " & CStr(Format(radius, "#0.000"))
        'Debug.Print Coords(1, i) & " " & Coords(2, i) & " " & Coords(3, i) & " " & Coords(4, i)
    Next i
        'Coords(4, UBound(Coords, 2) - 1) = Coords(4, UBound(Coords, 2) - 2)
    ArrSize = UBound(Coords, 2) + 2
    ReDim Preserve TunFile(ArrSize)

    For i = 1 To UBound(Coords, 2)
        spaceA = 26 - (Len(Coords(1, i)) + Len(Coords(2, i)))   'calculating number of space character so file formatting is readable for GEO
        spaceB = 38 - (Len(Coords(1, i)) + spaceA + Len(Coords(2, i)) + Len(Coords(3, i)))
        spaceC = 50 - (Len(Coords(1, i)) + spaceA + Len(Coords(2, i)) + spaceB + Len(Coords(3, i)) + Len(Coords(2, i)))
        spaceD = 61 - (Len(Coords(1, i)) + spaceA + Len(Coords(2, i)) + spaceB + Len(Coords(3, i)) + spaceC + Len(Coords(2, i)))
        spaceE = 75 - (Len(Coords(1, i)) + spaceA + Len(Coords(2, i)) + spaceB + Len(Coords(3, i)) + spaceC + Len(Coords(2, i)) + spaceD + Len(Coords(4, i)) + 1)
    
        TunFile(i + 2) = Coords(1, i) & Space(spaceA) & _
                         Coords(2, i) & Space(spaceB) & _
                         Coords(3, i) & Space(spaceC) & _
                         Coords(2, i) & Space(spaceD) & _
                         Coords(4, i) & Space(spaceE) & ","
                        
        'Debug.Print TunFile(i + 2)
    Next i
        TunFile(1) = RawTunFile(1)
        TunFile(2) = RawTunFile(2)
    'NewFilePath = Left(File_Path, InStrRev(File_Path, "\")) & Prefix & "_" & Format(mean, "#0.000") & ".tun"
    Set fileStream = fso.CreateTextFile(File_Path)
        
    For n = 1 To UBound(TunFile)
    fileStream.WriteLine TunFile(n)
    Next n
    fileStream.Close
    NoOfFiles = NoOfFiles + 1

Next counter

MsgBox ("All Done! Number of files with added radiuses = " & NoOfFiles)
End Sub

Sub List_Files_In_Directory()

Dim FileList() As Variant, i As Integer

FileList = Tools.File_List_In_Directory

For i = 1 To UBound(FileList)
    Debug.Print FileList(i)
Next i

End Sub

