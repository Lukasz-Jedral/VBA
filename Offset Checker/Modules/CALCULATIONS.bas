Attribute VB_Name = "CALCULATIONS"
Option Explicit
Option Base 1

Function CalculateAzimuth(CoordList() As String) As Double()

Dim i As Integer, j As Integer, k As Integer, n As Integer, cArrSize As Integer
Dim dX As Double, dY As Double, ArcTan As Double, Azimuth() As Double, Temp_List() As String
Dim StartPos As Integer, EndPos As Integer

j = 1

For i = 1 To UBound(CoordList, 2)
    If i = UBound(CoordList, 2) Then
    
        If CoordList(6, i) = "Closed" Then
            ReDim Preserve Temp_List(5, j)
            For k = 1 To 5
                Temp_List(k, j) = CoordList(k, i)
            Next k
        
            j = j + 1
        
            ReDim Preserve Temp_List(5, j)
        
            For n = 1 To UBound(CoordList, 2)
                If CoordList(5, i) = CoordList(5, n) Then
                    For k = 1 To 5
                        Temp_List(k, j) = CoordList(k, n)
                    Next k
                    Exit For
                End If
            Next n
        
            j = j + 1
        Else
            ReDim Preserve Temp_List(5, j)
            For k = 1 To 5
                Temp_List(k, j) = CoordList(k, i)
            Next k
            j = j + 1
        End If
    
    Else
        
        If CoordList(6, i) = "Closed" And CoordList(5, i) <> CoordList(5, i + 1) Then
    
            ReDim Preserve Temp_List(5, j)
            For k = 1 To 5
                Temp_List(k, j) = CoordList(k, i)
            Next k
        
            j = j + 1
        
            ReDim Preserve Temp_List(5, j)
        
            For n = 1 To UBound(CoordList, 2)
                If CoordList(5, i) = CoordList(5, n) Then
                    For k = 1 To 5
                        Temp_List(k, j) = CoordList(k, n)
                    Next k
                    Exit For
                End If
            Next n
        
            j = j + 1
        Else
            ReDim Preserve Temp_List(5, j)
            For k = 1 To 5
                Temp_List(k, j) = CoordList(k, i)
            Next k
            j = j + 1
        End If
    End If
Next i

cArrSize = 1

For k = 1 To UBound(Temp_List, 2) - 1
    If Temp_List(5, k + 1) = Temp_List(5, k) Then
    
        For i = k + 1 To UBound(Temp_List, 2)
            dX = Val(Temp_List(2, i)) - Val(Temp_List(2, i - 1))
            dY = Val(Temp_List(3, i)) - Val(Temp_List(3, i - 1))
    
            If dX <> 0 And dY <> 0 Then
                If dX > 0 And dY > 0 Then
                    ArcTan = Atn(Abs(dY) / Abs(dX))
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = ArcTan
                End If
    
                If dX < 0 And dY > 0 Then
                    ArcTan = Atn(Abs(dY) / Abs(dX))
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = 4 * Atn(1) - ArcTan
                End If
    
                If dX < 0 And dY < 0 Then
                    ArcTan = Atn(Abs(dY) / Abs(dX))
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = 4 * Atn(1) + ArcTan
                End If
    
                If dX > 0 And dY < 0 Then
                    ArcTan = Atn(Abs(dY) / Abs(dX))
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = 8 * Atn(1) - ArcTan
                End If
    
                If dX = 0 And dY > 0 Then
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = Atn(1)
                End If
    
                If dX = 0 And dY < 0 Then
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = 6 * Atn(1)
                End If
    
                If dX > 0 And dY = 0 Then
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = 8 * Atn(1)
                End If
    
                If dX < 0 And dY = 0 Then
                    ReDim Preserve Azimuth(cArrSize)
                    Azimuth(cArrSize) = 4 * Atn(1)
                End If
            
            cArrSize = cArrSize + 1
    
            Else
                Call MsgBox("At least two point in a line has the same coordinates." & vbNewLine & _
                "Impossible to calculate azimuth", vbExclamation, "Calcultion Error!")
                Exit Function
            End If
            Exit For
        Next i
    End If
Next k

CalculateAzimuth = Azimuth()

End Function

Sub Smallest_Offsets(Point_Coordinates() As String, Lines() As String, Azimuth() As Double, OrthoCoords() As String)

Dim dX As Double, dY As Double, Distance As Double, SideOffset As Double, TempOrthoCoords() As Double
Dim i As Integer, j As Integer, k As Integer, arrSize As Integer, LowestValue As Double, LowestValuePos As Integer, n As Integer



For i = 1 To UBound(Point_Coordinates, 2)
n = 1
    
    For j = 1 To UBound(Lines, 2)
        
        If j <> UBound(Lines, 2) Then
        
            If Lines(6, j) = "Open" And Lines(5, j) <> Lines(5, j + 1) Then
                j = j + 1
            End If
        
            dX = Val(Point_Coordinates(2, i)) - Val(Lines(2, j))
            dY = Val(Point_Coordinates(3, i)) - Val(Lines(3, j))
            Distance = dY * Sin(Azimuth(n)) + dX * Cos(Azimuth(n))
            SideOffset = dY * Cos(Azimuth(n)) - dX * Sin(Azimuth(n))
        
            ReDim Preserve TempOrthoCoords(2, n)
        
            TempOrthoCoords(1, n) = Distance
            TempOrthoCoords(2, n) = Abs(SideOffset)
        
            n = n + 1
        Else
            If Lines(6, j) = "Open" Then
                Exit For
            Else
                dX = Val(Point_Coordinates(2, i)) - Val(Lines(2, j))
                dY = Val(Point_Coordinates(3, i)) - Val(Lines(3, j))
                Distance = dY * Sin(Azimuth(n)) + dX * Cos(Azimuth(n))
                SideOffset = dY * Cos(Azimuth(n)) - dX * Sin(Azimuth(n))
        
                ReDim Preserve TempOrthoCoords(2, n)
        
                TempOrthoCoords(1, n) = Distance
                TempOrthoCoords(2, n) = Abs(SideOffset)
                
                n = n + 1
            End If
        End If
        
    Next j
    
    LowestValue = TempOrthoCoords(2, 1)
    LowestValuePos = 1
       
    For j = 1 To UBound(TempOrthoCoords, 2) - 1
        
        If LowestValue > TempOrthoCoords(2, j + 1) Then
           LowestValue = TempOrthoCoords(2, j + 1)
           LowestValuePos = j + 1
        End If
        
    Next j
    
        ReDim Preserve OrthoCoords(3, i)
        
    
        OrthoCoords(1, i) = Point_Coordinates(1, i)
        OrthoCoords(2, i) = TempOrthoCoords(1, LowestValuePos)
        OrthoCoords(3, i) = TempOrthoCoords(2, LowestValuePos)
Next i

  

End Sub

Sub Show_Points_Outside_Tolerance(Offsets() As String, ErrValue As Double)

Dim i As Integer, j As Integer
Dim Pts_To_Be_Checked() As String
Dim ArrCheck As Integer

j = 1

For i = 1 To UBound(Offsets, 2)
    If Offsets(3, i) >= ErrValue Then
        ReDim Preserve Pts_To_Be_Checked(3, j)
        Pts_To_Be_Checked(1, j) = Offsets(1, i)
        Pts_To_Be_Checked(2, j) = Offsets(2, i)
        Pts_To_Be_Checked(3, j) = Offsets(3, i)
        j = j + 1
    End If
Next i

On Error Resume Next

ArrCheck = UBound(Pts_To_Be_Checked, 2)

If Err.Number <> 9 Then

    Columns("A:A").ColumnWidth = 16.43
    Columns("B:B").ColumnWidth = 16.43
    Columns("C:C").ColumnWidth = 16.43

    Cells(1, 1) = "Points to be checked"
    Cells(2, 1) = "Point Name"
    Cells(2, 2) = "Distance along line"
    Cells(2, 3) = "Offset"

    For i = 1 To UBound(Pts_To_Be_Checked, 2)
            Cells(i + 2, 1) = Pts_To_Be_Checked(1, i)
        For j = 2 To 3
            Cells(i + 2, j) = CDbl(Format(Pts_To_Be_Checked(j, i), "###0.000"))
            Cells(i + 2, j).NumberFormat = "#,##0.000"
        Next j
    Next i
    
Call MsgBox("Points outside give tolerance are listed in active worksheet", vbOKOnly + vbExclamation, "Points to check!")
    
Else

Call MsgBox("All points withing given tolerance", vbOKOnly, "All good!")

End If

End Sub







































