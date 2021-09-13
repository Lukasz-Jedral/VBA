VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Charts_UserForm 
   Caption         =   "Create a chart"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8775
   OleObjectBlob   =   "Charts_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Charts_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox_Pt_No_Change()

End Sub

Private Sub UserForm_Initialize()

Dim nCol As Integer, nRow As Integer, i As Integer, j As Integer, sCol As Integer, sRow As Integer, n As Integer
'Dim sDate() As Integer, sNum() As Integer
Dim Ws As Worksheet, Arr() As String, ArrNext() As String
Dim Sekt() As String, Spont() As Integer, SpontNo As Integer, SektNo As Integer
    
Set Ws = Sheets("Spont")

'sDate = Charts.First_cell_with_data(Ws)
'sNum = Charts.First_cell_with_number(Ws)

ComboBox_Pt_No.Clear 'we start with clearing ComboBox in case there will be something left from previous operations
ComboBox_Date.Clear 'we start with clearing ComboBox in case there will be something left from previous operations
ComboBox_Section.Clear 'we start with clearing ComboBox in case there will be something left from previous operations

'--------------------------filling ComboBox with names of existing points assuming that point nubers are in first row----------------
nCol = Ws.Cells(3, Columns.Count).End(xlToLeft).Column 'nCol = column number of last filled cell

ComboBox_Pt_No.Style = fmStyleDropDownList

For i = 2 To nCol
    nRow = Ws.Cells(Rows.Count, i).End(xlUp).Row
    If IsNumeric(Ws.Cells(nRow, i)) = True Then
        With ComboBox_Pt_No
            .AddItem Cells(3, i)
        End With
    End If
Next

'--------------------------filling ComboBox with dates assuming that point nubers are in first column----------------
nRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 1st row 'finds the last non empty cell in row
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   Date Section   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
ComboBox_Date.Style = fmStyleDropDownList

For i = 4 To nRow
    With ComboBox_Date
        .AddItem Cells(i, 1)
        .List(i - 4) = Format(ComboBox_Date.List(i - 4), "Short Date")
    End With
Next

OptionButton_Data_One.Value = True
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   Section Section   >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

ComboBox_Section.Style = fmStyleDropDownList

For i = 3 To nCol
    sRow = Ws.Cells(Rows.Count, i).End(xlUp).Row
    If sRow > 3 Then
        sCol = i
        Exit For
    End If
Next i

For i = sCol To nCol - 1
    Arr = Split(Ws.Cells(3, i), ".")
    ArrNext = Split(Ws.Cells(3, i + 1), ".")
        If Arr(0) <> ArrNext(0) Then
           SpontNo = SpontNo + 1 '>>>>>makro counts how many times spont line number  changes. _
    >>>>>>>>>>>>>>>>>>>>>>>>>>>>For example 1.1 1.2 1.3 - the name changes 2 times from 1.1 to 1.2 and from 1.2 to 1.3, so if we want a nuber of section we must + 1
        End If
Next i
'SpontNo = SpontNo + 1 'now we know how many spont lines we have so we can resize Spont names Array
ReDim Spont(SpontNo)
For i = sCol To nCol - 1
    Arr = Split(Ws.Cells(3, i), ".")
    ArrNext = Split(Ws.Cells(3, i + 1), ".")
        If Arr(0) <> ArrNext(0) Then
           Spont(j) = Arr(0) 'we repeat the same procedure as above but this time when first part of sektion name changes we save the spont line number _
                            ex: previous point is 1.4.3 next is 2.1.1 so Arr(0)=1 and ArrNext(0)=2 we save into Spont(0) = 1
           j = j + 1
        End If
Next i

Spont(j) = Spont(j - 1) + 1 'here we save last spont line name. For example in exel file we chave points only in spont lines 2 and 3, so the name change will occur only once from 2 to 3. _
                            Because of that we will save in Spont() on position Spont(0) number 2 and we need also 3 on position Spont(1)
For j = 0 To UBound(Spont)
    For i = sCol To nCol - 1
        Arr = Split(Ws.Cells(3, i), ".")
        ArrNext = Split(Ws.Cells(3, i + 1), ".")
        If Arr(0) = Spont(j) And ArrNext(0) = Spont(j) And Arr(1) <> ArrNext(1) Then 'condicon: if Spont Line name is the same and sektion Number is changing then we count how many such changes there are
            SektNo = SektNo + 1
        End If
    Next i
Next j
SektNo = SektNo + 1 'number of changes + 1 gives us number of section
ReDim Sekt(SektNo + SpontNo - 1)

For j = 0 To UBound(Spont)
    For i = sCol To nCol - 1
        Arr = Split(Ws.Cells(3, i), ".")
        ArrNext = Split(Ws.Cells(3, i + 1), ".")
        If Arr(0) = Spont(j) And ArrNext(0) = Spont(j) And Arr(1) <> ArrNext(1) Then 'condicon: if Spont Line name is the same and sektion Number is changing then we count how many such changes there are
            Sekt(n) = Spont(j) & "." & Arr(1) 'afer condition is meet we save section name into array ex: 1.3 to 1.4 we will save 1.3
            n = n + 1
        End If
        If Arr(0) = Spont(j) And ArrNext(0) <> Spont(j) Then 'previous condition wasn't taking into accauny situation in which also spont line nuber changes for ex. 1.4 --> 2.1
            Sekt(n) = Spont(j) & "." & Arr(1)
            n = n + 1
        End If
    Next i
Next j
Sekt(n) = Spont(j - 1) & "." & Arr(1) 'adding last sektion name to atrray. Function is based on changes so in case of last change we need this extra line. For example last two sektion are 3.2 and 3.3 _
                                            so last change is from 3.2 to 3.3 and macro will save name of 3.2 sektion. 3.3 will be not saved because there is no 3.4 sektion.
For i = 0 To UBound(Sekt)
    With ComboBox_Section
        .AddItem Sekt(i)
    End With
Next i
End Sub

Private Sub CreateChartPointButton_Click()

Call Charts.Chart_for_point

End Sub
Private Sub CreateChartSectionButton_Click()

Call Charts.Chart_for_section

End Sub

Private Sub CreateChartDateButton_Click()

Call Charts.Chart_for_date

End Sub

Private Sub Delete_Charts_Button_Click()
Dim wks As Worksheet, ans As Integer

On Error Resume Next

Application.DisplayAlerts = False

ans = MsgBox("Delete all charts?", vbQuestion + vbYesNo)

If ans = vbYes Then

For Each wks In Worksheets
    If wks.ChartObjects.Count > 0 And Not wks.Name = "Spont" Then
        wks.Delete
    Else
        wks.ChartObjects.Delete
    End If
Next wks

ActiveWorkbook.Charts.Delete

End If

Application.DisplayAlerts = True

'Me.Hide

Sheets("Spont").Activate

On Error GoTo 0

End Sub

Private Sub ExitButton_Click()

'Me.Hide
Unload Me

End Sub
