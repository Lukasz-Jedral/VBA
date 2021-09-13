Attribute VB_Name = "Displacments"
Option Explicit
Option Base 1
Public msg As String, meas_Date As String
Sub A_Run_Macro()
LoadDataUserForm.Show
End Sub
Function GetTheFilePath()
'function will show pop up window for user to choose file and returns path to that file

'defining variables
Dim File_Path As String
Dim UserID As String
Dim intChoice As Integer
'UserID = Application.UserName 'getting current user name

'----------------------------------------- seting up starting directory for browing files depending on user name--------------------------
'----------------------------------------- if its not me or Mathias sysem default will apply----------------------------------------------
'If UserID = "Jedral, Lukasz" Then
    Application.FileDialog(msoFileDialogOpen _
    ).InitialFileName = "C:\Users\JedralL\SharePoint\Solna United - 00. Delade dokument\Väg och Anl Sthlm\11 Mätning\03 - Inmätning\Kontrollmätning\Spont\02 - Mätdata"
'End If

'If UserID = "Kartano, Mathias" Then
    'Application.FileDialog(msoFileDialogOpen _
    ').InitialFileName = "G:\Väg och Anl Sthlm\Gemensam\71340-DC Daniel Garpsäter\71472 PrC Maryam Zarrin\Solna United, kv. Tygeln\11 Mätning\03 - Inmätning\Kontrollmätning\Områdespåverkan"
'End If
'------------------------GET THE FILE NAME WHICH WILL BE OPENED----------------------

'only allow the user to select one file
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
'change the display name of the open file dialog
Application.FileDialog(msoFileDialogOpen).Title = "Select Coordinate Document"
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
Function ExtractCoordinatesFromGeoFile(ByRef Coords() As String, File_Path As String)
'function extract coordinates from .geo format file and returns those coordinates as array

Dim LineFromFile As String, LineItems As String
Dim Matrix() As String, TempSplited() As String
Dim i As Integer, j As Integer, k As Integer, StartPos As String, EndPos As String

If File_Path = "" Then
Call MsgBox("Incorrect file path. Please load correct file and try again", vbCritical + vbOKCancel, "That will not work!")
Exit Function
Else
Open File_Path For Input As #1 'opens file as file no. 1 in "to be read" mode
End If

Do Until EOF(1) 'loop until end of file

Line Input #1, LineFromFile 'reads line from file no. 1

i = i + 1
ReDim Preserve Matrix(i)
Matrix(i) = LineFromFile
Matrix(i) = Replace(Matrix(i), vbTab, "")

Loop
Close #1
'---------------------CHECK IF THE FILE FORMAT IS OK---------------------------------
If Matrix(1) <> "FileHeader " & Chr(34) & "SBG Object Text v2.01" & Chr(34) & "," & Chr(34) & "Coordinate Document" & Chr(34) Then
Call MsgBox("Unsupported file format", vbExclamation, "Warning!") 'add instructions what caused erron and how should it be handled!!!!!!!!!!!!!!!!!!!!!!!!!!!
Exit Function
End If
j = 1
For i = 1 To UBound(Matrix())
    If InStr(1, Matrix(i), "Point ", vbTextCompare) <> 0 Then
        TempSplited = Split(Matrix(i), ",", , vbBinaryCompare)
        ReDim Preserve Coords(4, j)
            For k = 1 To 4
            Coords(k, j) = TempSplited(k - 1)
            Next k
        StartPos = InStr(1, Coords(1, j), Chr(34), vbBinaryCompare)
        EndPos = InStrRev(Coords(1, j), Chr(34), , vbBinaryCompare)
        Coords(1, j) = Mid(Coords(1, j), StartPos + 1, EndPos - StartPos - 1)
        j = j + 1
    End If
Next i
End Function
Sub Caculations()
'procedure calculates offsets of each point in input file against given reference line. Zero measurent, current measurent and reference line is requaierd as imput

Dim i As Integer, j As Integer, Pt_Count As Integer, nRow As Integer, nCol As Integer, icount As Integer
Dim dX As Double, dY As Double, ArcTan As Double, Azimuth() As Double, dH As Double
Dim Temp_Ref_Point(1 To 4) As String, Temp_Meas_Point(1 To 4) As String, Temp_Azimuth() As String, Temp_Ref_Line_Start(1 To 4) As String
Dim Ortho_Ref_Point() As String, Ortho_Meas_Point() As String, Results() As String, TempResults() As String
Dim Ref_Line_File_Dir As String, Ref_Points_File_Dir As String, Meas_Points_Dir As String
Dim Ref_Line_File() As String, Ref_Points_File() As String, Meas_Points() As String
Dim Ref_Line_No As Byte, meas_Date As Date

Dim Dist As Double
'every array format is |PtNo X Y Z| _
                       |PtNo X Y Z| _
                       |PtNo X Y Z| and so on...

If Dir(ActiveWorkbook.Path & "\Excel_Macro_Data\RefLineDir.txt", vbDirectory) = vbNullString Then
Call MsgBox("Please load file contaning Reference Line Coordinetes before continuing.", vbCritical + vbOKOnly, "Reference Lines Missing")
Exit Sub
End If

If Dir(ActiveWorkbook.Path & "\Excel_Macro_Data\RefPointsDir.txt", vbDirectory) = vbNullString Then
Call MsgBox("Please load file contaning Reference Points Coordinetes before continuing.", vbCritical + vbOKOnly, "Reference Points Missing")
Exit Sub
End If

Open ActiveWorkbook.Path & "\Excel_Macro_Data\RefLineDir.txt" For Input As #1
    Line Input #1, Ref_Line_File_Dir
Close #1
    Ref_Line_File_Dir = Mid(Ref_Line_File_Dir, 2, Len(Ref_Line_File_Dir) - 2)
    
Open ActiveWorkbook.Path & "\Excel_Macro_Data\RefPointsDir.txt" For Input As #1
    Line Input #1, Ref_Points_File_Dir
Close #1
    Ref_Points_File_Dir = Mid(Ref_Points_File_Dir, 2, Len(Ref_Points_File_Dir) - 2)
    
meas_Date = Application.InputBox("Type in measurement date!(Date format 2016-01-31)", "Type in date", FormatDateTime(Date, vbShortDate), Type:=1)
If meas_Date = "00:00:00" Or IsDate(meas_Date) = False Then
Call MsgBox("Given value is not a date", vbOKOnly, "Incorect date format")
Exit Sub
End If

Meas_Points_Dir = Displacments.GetTheFilePath 'this should be pointed every time

If Ref_Line_File_Dir = "" Or Ref_Points_File_Dir = "" Then
    Call MsgBox("Incorrect file path. Please load correct file and try again", vbCritical + vbOKOnly, "That will not work!")
    Exit Sub
End If

Application.ScreenUpdating = False

Call Displacments.ExtractCoordinatesFromGeoFile(Ref_Line_File(), Ref_Line_File_Dir)
Call Displacments.ExtractCoordinatesFromGeoFile(Ref_Points_File(), Ref_Points_File_Dir)
Call Displacments.ExtractCoordinatesFromGeoFile(Meas_Points(), Meas_Points_Dir)

On Error GoTo ErrHandler:
    If UBound(Ref_Line_File) = 1 Or UBound(Ref_Points_File()) = 1 Or UBound(Meas_Points()) = 1 Then
    End If
    
ErrHandler:
    If Err.Number = 9 Then
        Call MsgBox("Incorrect file path. Please load correct file and try again", vbCritical + vbOKOnly, "That will not work!")
        Exit Sub
    End If
    
    
    
nCol = Cells(3, Columns.Count).End(xlToLeft).Column 'finds the last non empty cell in 1st column
nRow = Cells(Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 1st row

Cells(nRow + 1, 1) = meas_Date 'put the data to the header of new row

'---first step we pick a Ref_Point and we check to which Ref Line it is closest to---
'to do that we have to compute azimuts of every ref line ------> Fuction CalculateAzimuth
Azimuth() = Displacments.CalculateAzimuth(Ref_Line_File)

For i = 1 To UBound(Ref_Points_File, 2)
    For j = 1 To 4
        Temp_Ref_Point(j) = Ref_Points_File(j, i) 'we pick one of the ref points and check to which ref line it is closest to
    Next j
    Ref_Line_No = Displacments.Find_The_Closest_Ref_Line(Temp_Ref_Point(), Ref_Line_File(), Azimuth())
        If Ref_Line_No = 0 Then
            GoTo NextIteration
        End If
    'now we know to which reference line the ref point is closest to. The next step is to find all measured points that are close to that ref point
    For Pt_Count = 1 To UBound(Meas_Points(), 2)
       Dist = Sqr(((Val(Temp_Ref_Point(2)) - Val(Meas_Points(2, Pt_Count))) ^ 2) + ((Val(Temp_Ref_Point(3)) - Val(Meas_Points(3, Pt_Count))) ^ 2))
        'Dist = Sqr(|dX ^ 2 + dY ^ 2|)
       dH = Val(Temp_Ref_Point(4)) - Val(Meas_Points(4, Pt_Count))
    If Dist < 1 And Abs(dH) < 0.5 Then 'if the horizontal distace between ref and meas point is less than 1 m and hihgt diff less than 0,5 m then Orthogonal Distance and Side Offset is calculated for both points and _
                        Offsets are compared - this is the value which will be shown. To do that macro will:
               
        For j = 1 To 4
        Temp_Ref_Line_Start(j) = Ref_Line_File(j, Ref_Line_No) 'extract coords of starting point of ref line we are reffering to
        Next j
        
        For j = 1 To 4
        Temp_Meas_Point(j) = Meas_Points(j, Pt_Count) 'extract coords of measured point
        Next j
        
        Ortho_Ref_Point() = Displacments.Get_Ortho_Dist_Offset(Temp_Ref_Point(), Temp_Ref_Line_Start(), Azimuth(Ref_Line_No)) 'calculate offset for ref point
        Ortho_Meas_Point() = Displacments.Get_Ortho_Dist_Offset(Temp_Meas_Point(), Temp_Ref_Line_Start(), Azimuth(Ref_Line_No)) 'calculate offset for meas point
        
        '--------------------Filling the value into right cell in table--------------------------
        'the row we know. This is next after the last that was fillend in the moment of running macro so nRow + 1
        'the only thing is to find the name of ref point which is closest to measured point in the table
        
        
        For icount = 1 To nCol
            If Temp_Ref_Point(1) = Cells(3, icount) Then
            Cells(nRow + 1, icount) = CDbl(Format(Ortho_Meas_Point(3) - Ortho_Ref_Point(3), "###0.000"))
            Cells(nRow + 1, icount).NumberFormat = "0.000"
            End If
        Next icount
        
        'Cells(i + 1, nCol + 1) = CDbl(Format(Ortho_Meas_Point(3) - Ortho_Ref_Point(3), "###0.000"))
        'Cells(i + 1, nCol + 1).NumberFormat = "0.000"
              
        On Error Resume Next
        If UBound(Results, 2) = 2 Then 'this condition will be fullfiled for first 2 collums on Result matrix _
                                        at first step we will have empty matrix then Ubound will result with Error - so resuming next _
                                        with second collumn condtion will be fullfiled so also lines of code for this if procedure will be executed
            ReDim Preserve Results(2, 2) 'set the size of the matrix results is Results(2,2) - so already has 2 collumn and 2 rows
            
            If Results(1, 1) = "" Then
                Results(1, 1) = Ortho_Ref_Point(1)         'this and next line - filling first collumn with data
                Results(2, 1) = Ortho_Meas_Point(3) - Ortho_Ref_Point(3)
            Else                                         'we will start with else when matrix will have first collumn filled
                Results(1, 2) = Ortho_Ref_Point(1)       'this and next line - filling second collumn with data
                Results(2, 2) = Ortho_Meas_Point(3) - Ortho_Ref_Point(3)
                ReDim Preserve Results(2, UBound(Results, 2) + 1) 'add empty collumn at end of matrix. This collumn will be filled at next loop step
            End If
        Else 'now with Results matrix which have 3 and more collumns code bellow will be executed
            Results(1, UBound(Results, 2)) = Ortho_Ref_Point(1)       'this and next line - filling added in previous loop step collumn with data
            Results(2, UBound(Results, 2)) = Ortho_Meas_Point(3) - Ortho_Ref_Point(3)
            ReDim Preserve Results(2, UBound(Results, 2) + 1) 'add next collumn to the matrix results - Ubound returs how many collumns current matrix have
        End If

    End If
    Next Pt_Count
NextIteration:
Next i
ReDim Preserve Results(2, UBound(Results, 2) - 1) 'removing last collumn from matrix. _
                                                   It is empty because extra collum was added after filling matrix with last data
Application.ScreenUpdating = True
For i = 1 To 2
    For j = 1 To UBound(Results, 2)
        Debug.Print Format(Results(i, j), "# ###.00000") 'show in immediate window Result Matrix
    Next j
Next i

Call Show_And_Or_Save_Raport(Results(), meas_Date)

End Sub


Function CalculateAzimuth(CoordList() As String) As Double()
'function calculates azimuth of each line in polyline

Dim i As Integer, j As Integer
Dim dX As Double, dY As Double, ArcTan As Double, Azimuth() As Double

For i = 2 To UBound(CoordList, 2)
    dX = Val(CoordList(2, i)) - Val(CoordList(2, i - 1))
    dY = Val(CoordList(3, i)) - Val(CoordList(3, i - 1))
    
If dX <> 0 And dY <> 0 Then
    If dX > 0 And dY > 0 Then
    ArcTan = Atn(Abs(dY) / Abs(dX))
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = ArcTan
    End If
    
    If dX < 0 And dY > 0 Then
    ArcTan = Atn(Abs(dY) / Abs(dX))
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = 4 * Atn(1) - ArcTan
    End If
    
    If dX < 0 And dY < 0 Then
    ArcTan = Atn(Abs(dY) / Abs(dX))
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = 4 * Atn(1) + ArcTan
    End If
    
    If dX > 0 And dY < 0 Then
    ArcTan = Atn(Abs(dY) / Abs(dX))
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = 8 * Atn(1) - ArcTan
    End If
    
    If dX = 0 And dY > 0 Then
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = Atn(1)
    End If
    
    If dX = 0 And dY < 0 Then
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = 6 * Atn(1)
    End If
    
    If dX > 0 And dY = 0 Then
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = 8 * Atn(1)
    End If
    
    If dX < 0 And dY = 0 Then
    ReDim Preserve Azimuth(i - 1)
    Azimuth(i - 1) = 4 * Atn(1)
    End If
    
Else
Call MsgBox("At least two point in a line has the same coordinates." & vbNewLine & _
"Impossible to calculate azimuth", vbExclamation, "Calcultion Error!")
Exit Function
End If

Next i

CalculateAzimuth = Azimuth()

End Function

Function Find_The_Closest_Ref_Line(Point_Coord() As String, Ref_Line_Coords() As String, Azimuth() As Double) As Byte
'Function determines which reference line is closest to given point.

Dim dX As Double, dY As Double, Distance As Double, SideOffset As Double, OrthoCoords() As Double
Dim i As Integer, j As Integer

Find_The_Closest_Ref_Line = 0

For i = 1 To UBound(Azimuth())

    dX = Val(Point_Coord(2)) - Val(Ref_Line_Coords(2, i))
    dY = Val(Point_Coord(3)) - Val(Ref_Line_Coords(3, i))
    Distance = dY * Sin(Azimuth(i)) + dX * Cos(Azimuth(i))
    If Distance < 0 Then
        SideOffset = 1000
    Else
    SideOffset = dY * Cos(Azimuth(i)) - dX * Sin(Azimuth(i))
    End If
        ReDim Preserve OrthoCoords(i)
        OrthoCoords(i) = Abs(SideOffset)
Next i


If Application.Min(OrthoCoords()) > 1.2 Then
Find_The_Closest_Ref_Line = 0
Exit Function
Else
    'For i = 1 To UBound(OrthoCoords)
        'If OrthoCoords(i) < 0 Then
            'OrthoCoords(i) = Abs(OrthoCoords(i) * 1000)
        'End If
    'Next i
    
    Find_The_Closest_Ref_Line = Application.Match(Application.Min(OrthoCoords()), OrthoCoords(), 0) 'finds the position of the minimal value of _
                                                                                                  SideOffset which refers to Ref Line Number
End If

End Function

Function Get_Ortho_Dist_Offset(Point_Coord() As String, Ref_Line_Coords() As String, Azimuth As Double) As String()
'Function determines reference line orthogonal coordinates of given points

Dim dX As Double, dY As Double, Distance As Double, SideOffset As Double
Dim i As Integer, j As Integer
Dim Ortho() As String

    dX = Val(Point_Coord(2)) - Val(Ref_Line_Coords(2))
    dY = Val(Point_Coord(3)) - Val(Ref_Line_Coords(3))
    Distance = dY * Sin(Azimuth) + dX * Cos(Azimuth)
    SideOffset = dY * Cos(Azimuth) - dX * Sin(Azimuth)
    
    ReDim Ortho(1 To 3)
    
    Ortho(1) = Point_Coord(1)
    Ortho(2) = Distance
    Ortho(3) = SideOffset

Get_Ortho_Dist_Offset = Ortho

End Function

Sub SaveRefFileDirectory(i As Integer)

Dim File_Path As String, Split_Dir() As String

If Dir(ActiveWorkbook.Path & "\Excel_Macro_Data", vbDirectory) = vbNullString Then
MkDir (ActiveWorkbook.Path & "\Excel_Macro_Data")
End If


File_Path = Displacments.GetTheFilePath

If i = 1 Then
    Open ActiveWorkbook.Path & "\Excel_Macro_Data\RefLineDir.txt" For Output As #1
    Write #1, File_Path
    Close #1
Else
    Open ActiveWorkbook.Path & "\Excel_Macro_Data\RefPointsDir.txt" For Output As #1
    Write #1, File_Path
    Close #1
End If

'Split_Dir = Split(File_Path, "\")
'Call MsgBox("File" & " " & Split_Dir(UBound(Split_Dir)) & " " & "was loaded", vbOKOnly + vbInformation)

End Sub

Sub Show_And_Or_Save_Raport(myArray() As String, meas_Date As Date)

Dim i As Long

msg = "Measurment date:" & " " & meas_Date & vbCrLf

For i = LBound(myArray, 2) To UBound(myArray, 2)
    msg = msg & myArray(1, i) & "     " & Format(myArray(2, i), "###0.000") & vbCrLf
    Next i
  
UserFromRaport.Show
    
End Sub














