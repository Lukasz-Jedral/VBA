Attribute VB_Name = "Execute"
Option Explicit
Option Base 1

Sub Execute()

Dim ErrValue As Double
Dim File_Path As String
Dim Points() As String, Lines() As String
Dim Azimuths() As Double, Offsets() As String

ErrValue = Application.InputBox("Please enter acceptable offset limit", _
"Offset limit value", "0,02", , , , , 1) 'setting the acceptable value of offset limit

If ErrValue = 0 Then
    Call MsgBox("Makro canceled by user!", vbOKOnly + vbCritical)
    Exit Sub
End If

File_Path = INPUT_DATA.GetTheFilePath

If File_Path = "" Then
    Call MsgBox("No file loaded!", vbOKOnly + vbCritical)
    Exit Sub
End If

ActiveSheet.Cells.ClearContents

Points() = INPUT_DATA.Import_Points(File_Path)

Lines() = INPUT_DATA.Import_Lines(File_Path)

On Error Resume Next

If UBound(Lines) <> 0 Or UBound(Points) <> 0 Then
    Else
    Call MsgBox("Missing Points or Lines. Please check input data", vbOKOnly + vbCritical)
    Exit Sub
End If

On Error GoTo 0

Azimuths() = CALCULATIONS.CalculateAzimuth(Lines())

Call CALCULATIONS.Smallest_Offsets(Points(), Lines(), Azimuths(), Offsets())

Call CALCULATIONS.Show_Points_Outside_Tolerance(Offsets(), ErrValue)

End Sub
