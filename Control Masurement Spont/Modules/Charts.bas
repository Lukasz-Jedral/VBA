Attribute VB_Name = "Charts"
Option Explicit

Sub Chart_for_point()


Dim nRow As Integer, nCol As Integer, i As Integer, sCol As Integer, sRow As Integer, txt As String
Dim Ws As Worksheet, wsLarm As Worksheet
'Dim sDate() As Integer, sNum() As Integer
Dim myVal() As Variant, myXVal(1) As Double, Arr() As Double
Dim SheetExist As Boolean, YesNo As Integer
Dim Vals() As Variant, LogicTest As Boolean, j As Integer

Set Ws = Sheets("Spont")

SheetExist = Charts.sheetExists(Charts_UserForm.ComboBox_Pt_No.text)

If SheetExist = True Then
    YesNo = MsgBox("Chart for Point: " & Charts_UserForm.ComboBox_Pt_No.text & " already exist" & vbNewLine & "To proceed this chart has to be deleted." & vbNewLine & "Delete chart?", vbExclamation + vbYesNo)
    If YesNo = vbYes Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ActiveWorkbook.Sheets(Charts_UserForm.ComboBox_Pt_No.text).Delete
        Application.DisplayAlerts = True
    Else
        Call MsgBox("Existing chart was keept." & vbNewLine & "Closing macro.", vbInformation + vbOKOnly)
        Exit Sub
    End If
End If


'sDate = Charts.First_cell_with_data(Ws)
'sNum = Charts.First_cell_with_number(Ws)

nCol = Ws.Cells(3, Ws.Columns.Count).End(xlToLeft).Column 'finds the last non empty cell in 1st column
nRow = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 1st row

'Call Create_Alarm_Levels(nCol)

'Set wsLarm = Sheets("AlarmLevels")

For i = 2 To nCol
    If Ws.Cells(3, i).text = Charts_UserForm.ComboBox_Pt_No.Value Then
       sCol = i
    End If
Next

For i = 4 To nRow
    If Ws.Cells(i, sCol) <> vbNullString Then
        sRow = i
        Exit For
    End If
Next i

myVal = Array(0, 1)
Arr = Charts.Alarm_Levels(Charts_UserForm.ComboBox_Pt_No.text)

ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count), Type:=xlChart
ActiveWorkbook.ActiveSheet.Name = Charts_UserForm.ComboBox_Pt_No.text

With ActiveChart
    .ChartType = xlBarClustered
    .SetSourceData Source:=Ws.Range(Ws.Cells(sRow, sCol), Sheet1.Cells(nRow, sCol))
    .SeriesCollection(1).XValues = Ws.Range(Ws.Cells(sRow, 1), Sheet1.Cells(nRow, 1))
    .HasTitle = True
    .ChartTitle.text = "Spont movement chart" & " - " & "Point Number:" & " " & Charts_UserForm.ComboBox_Pt_No.Value
    .Axes(xlCategory).TickLabelPosition = xlLow
    .SeriesCollection(1).Name = "Measured data"
    '.Axes(xlValue).MinimumScale = -0.03
    '.Axes(xlValue).MaximumScale = 0.03
    .Axes(xlValue).MajorUnit = 0.01
    .SetElement (msoElementPrimaryValueGridLinesMinorMajor)
    .Axes(xlValue).MinorUnit = 0.001
    .Axes(xlCategory).CategoryType = xlCategoryScale
'---------------adding alarm lines, remember to set conditions when which lvl are applied---------
For i = 1 To 2
    .SeriesCollection.NewSeries
    .SeriesCollection(i + 1).Name = "Alarm Level " & i
    .SeriesCollection(i + 1).ChartType = xlXYScatterLinesNoMarkers
    .SeriesCollection(i + 1).Values = myVal
    myXVal(0) = Arr(i - 1)
    myXVal(1) = Arr(i - 1)
    .SeriesCollection(i + 1).XValues = myXVal
    .Axes(xlValue, xlSecondary).MaximumScale = 1
    .Axes(xlValue, xlSecondary).MinimumScale = 0
    .SeriesCollection(i + 1).Format.Line.Visible = msoTrue
    .SeriesCollection(i + 1).Format.Line.DashStyle = msoLineLongDash
    If i = 1 Then
        .SeriesCollection(i + 1).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent3
    Else
        .SeriesCollection(i + 1).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
    End If
Next i
    .Axes(xlValue, xlSecondary).Delete
End With

Charts_UserForm.Hide

End Sub
Sub Chart_for_section()
Dim Ws As Worksheet
Dim i As Integer, j As Integer, nRow As Integer, nCol As Integer, nCount As Integer, sCol As Integer, sRow As Integer
Dim Arr() As String, Sektion As String, Vals() As Variant, LogicTest As Integer
Dim SheetExist As Boolean, YesNo As Integer
    
Set Ws = Excel.ActiveWorkbook.Sheets("Spont")

SheetExist = Charts.sheetExists("Sektion " & Charts_UserForm.ComboBox_Section.text)

If SheetExist = True Then
    YesNo = MsgBox("Chart for Sektion " & Charts_UserForm.ComboBox_Section & " already exist" & vbNewLine & "To proceed this chart has to be deleted." & vbNewLine & "Delete chart?", vbExclamation + vbYesNo)
    If YesNo = vbYes Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ActiveWorkbook.Sheets("Sektion " & Charts_UserForm.ComboBox_Section.text).Delete
        Application.DisplayAlerts = True
    Else
        Call MsgBox("Existing chart was keept." & vbNewLine & "Closing macro.", vbInformation + vbOKOnly)
        Exit Sub
    End If
End If

nCol = Ws.Cells(3, Ws.Columns.Count).End(xlToLeft).Column 'finds the last non empty cell in 1st column
nRow = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 1st row

For i = 2 To nCol
    Arr = Split(Ws.Cells(3, i), ".")
    Sektion = Arr(0) & "." & Arr(1)
        If Sektion = Charts_UserForm.ComboBox_Section.Value Then
            sCol = i
            Exit For
        End If
Next i

For i = 2 To nCol
    Arr = Split(Ws.Cells(3, i), ".")
    Sektion = Arr(0) & "." & Arr(1)
        If Sektion = Charts_UserForm.ComboBox_Section.Value Then
               nCount = nCount + 1
        End If
Next i

nCount = nCount - 1
sRow = 999

For i = sCol To sCol + nCount
    For j = 4 To nCol
        If IsNumeric(Ws.Cells(j, i)) = True And IsEmpty(Ws.Cells(j, i)) = False Then
            If j < sRow Then
                sRow = j
            End If
            Exit For
        End If
    Next j
        
Next i

ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count), Type:=xlChart
ActiveChart.Name = "Sektion " & Charts_UserForm.ComboBox_Section.Value

    With ActiveChart
        .ChartType = xlLine
        .ChartStyle = 232
        .PlotArea.Format.Fill.Visible = msoFalse
        .SetSourceData Source:=Ws.Range(Ws.Cells(sRow, sCol), Ws.Cells(nRow, sCol + nCount))
        .SeriesCollection(1).XValues = Ws.Range(Ws.Cells(sRow, 1), Sheet1.Cells(nRow, 1))
        .HasTitle = True
        .ChartTitle.text = "Spont movement chart" & " - " & "Section:" & " " & Charts_UserForm.ComboBox_Section.Value
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlCategory).TickLabels.Orientation = xlDownward
        .Axes(xlCategory).CategoryType = xlCategoryScale
    For i = .SeriesCollection.Count To 1 Step -1
        .SeriesCollection(i).ChartType = xlLineMarkers
        .SeriesCollection(i).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(i).Format.Fill.Visible = msoTrue
        '.SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .SeriesCollection(i).Name = "Level " & i
        '.SeriesCollection(i).MarkerSize = 10
        
        Vals = .SeriesCollection(i).Values
        For j = 1 To .SeriesCollection(i).Points.Count
            If Not Vals(j) = vbNullString Then
                LogicTest = 1
            End If
        Next j
        
        If LogicTest = 0 Then
            .Legend.LegendEntries(i).Delete
        End If
    Next i
    
    End With
Call Charts.ChangePointColor_DataCharts_Sections(sCol)
    

End Sub

Sub Chart_for_date()

If Charts_UserForm.OptionButton_Data_All.Value = True Then

    Call Charts.Chart_for_date_all

Else

    Call Charts.Chart_for_date_one

End If

End Sub

Sub Chart_for_date_all()

Dim WsChrt As Worksheet, Ws As Worksheet
Dim chrt As Shape
Dim nCol As Integer, nRow As Integer, sRow As Integer, i As Integer, j As Integer, iTop As Integer, n As Integer, IfNotEmpty As Integer
Dim myVal() As Double, Arr() As Double, myXVal() As Double
Dim Sekt_Names() As String, Data_Series() As Variant, Series1() As Variant, Series2() As Variant, Series3() As Variant
Dim Series As String, Vals() As Variant, LogicTest As Integer


Application.DisplayAlerts = False
On Error Resume Next
ActiveWorkbook.Sheets("Charts for date").Delete
Application.DisplayAlerts = True

ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count), Type:=xlWorksheet
Sheets(Sheets.Count).Name = "Charts for date"

Set WsChrt = ActiveSheet
Set Ws = ActiveWorkbook.Sheets("Spont")

'sDate = Charts.First_cell_with_data(Ws)
'sNum = Charts.First_cell_with_number(Ws)

nCol = Ws.Cells(3, Ws.Columns.Count).End(xlToLeft).Column 'finds the last non empty cell in 1st column
nRow = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 3st row

'Call Create_Alarm_Levels(nCol)
  

For sRow = 4 To nRow Step 2

Sekt_Names = Charts.Create_Sektion_Names(nCol, sRow)
Data_Series = Charts.Create_Data_For_Series(Sekt_Names, sRow, nCol, nRow, 3)

ReDim Series1(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series1(i) = Data_Series(0, i)
Next i

ReDim Series2(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series2(i) = Data_Series(1, i)
Next i

ReDim Series3(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series3(i) = Data_Series(2, i)
Next i

Set chrt = WsChrt.Shapes.AddChart2(, xlLineMarkers, 4, iTop, 421, 297)
    With chrt
        .Name = "Chart" & sRow
    End With

    With chrt.Chart
        .ChartType = xlLine
        .ChartStyle = 232
        .PlotArea.Format.Fill.Visible = msoFalse
        '.SetSourceData Source:=Ws.Range(Ws.Cells(sRow, 2), Ws.Cells(sRow, nCol))
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = Series1
        .SeriesCollection(1).XValues = Sekt_Names() 'Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, nCol))
        .HasTitle = True
        .ChartTitle.text = "Spont movement chart" & " - " & "Measurent Date:" & " " & Ws.Cells(sRow, 1).Value
        .Axes(xlCategory).TickLabelPosition = xlLow
        .SeriesCollection(1).Values = Series2
        '.SeriesCollection(1).Name = "Spont Level 1"
        .SeriesCollection(1).ChartType = xlLineMarkers
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).Format.Fill.Visible = msoTrue
        .SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = Series1
        .SeriesCollection(2).Name = "Spont Level 1"
        .SeriesCollection(2).XValues = Sekt_Names()
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = Series2
        .SeriesCollection(3).Name = "Spont Level 2"
        .SeriesCollection(3).XValues = Sekt_Names()
        .SeriesCollection.NewSeries
        .SeriesCollection(4).Values = Series3
        .SeriesCollection(4).Name = "Spont Level 3"
        .SeriesCollection(4).XValues = Sekt_Names()
        .SetElement (msoElementLegendBottom)
        '.Legend.LegendEntries(1).Delete
        
        For i = .SeriesCollection.Count To 2 Step -1
            .SeriesCollection(i).ChartType = xlLineMarkers
            .SeriesCollection(i).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(i).Format.Fill.Visible = msoTrue
            '.SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
            '.SeriesCollection(i).Name = "Level " & i - 1
            '.SeriesCollection(i).MarkerSize = 10
            LogicTest = 0
            
        Vals = .SeriesCollection(i).Values
            For j = 1 To .SeriesCollection(i).Points.Count
                If Not Vals(j) = vbNullString Then
                    LogicTest = 1
                End If
            Next j
        
            If LogicTest = 0 Then
                .Legend.LegendEntries(i).Delete
            End If
        Next i
         .Legend.LegendEntries(1).Delete
    End With

iTop = iTop + 306 '297

Next sRow

iTop = 0
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
For sRow = 5 To nRow Step 2

Sekt_Names = Charts.Create_Sektion_Names(nCol, sRow)
Data_Series = Charts.Create_Data_For_Series(Sekt_Names, sRow, nCol, nRow, 3)

ReDim Series1(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series1(i) = Data_Series(0, i)
Next i

ReDim Series2(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series2(i) = Data_Series(1, i)
Next i

ReDim Series3(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series3(i) = Data_Series(2, i)
Next i

Set chrt = WsChrt.Shapes.AddChart2(, xlLineMarkers, 429, iTop, 421, 297)
    With chrt
        .Name = "Chart" & sRow
    End With

    With chrt.Chart
        .ChartType = xlLine
        .ChartStyle = 232
        .PlotArea.Format.Fill.Visible = msoFalse
        '.SetSourceData Source:=Ws.Range(Ws.Cells(sRow, 2), Ws.Cells(sRow, nCol))
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = Series1
        .SeriesCollection(1).XValues = Sekt_Names() 'Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, nCol))
        .HasTitle = True
        .ChartTitle.text = "Spont movement chart" & " - " & "Measurent Date:" & " " & Ws.Cells(sRow, 1).Value
        .Axes(xlCategory).TickLabelPosition = xlLow
        .SeriesCollection(1).Values = Series2
        '.SeriesCollection(1).Name = "Spont Level 1"
        .SeriesCollection(1).ChartType = xlLineMarkers
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).Format.Fill.Visible = msoTrue
        .SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = Series1
        .SeriesCollection(2).Name = "Spont Level 1"
        .SeriesCollection(2).XValues = Sekt_Names()
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = Series2
        .SeriesCollection(3).Name = "Spont Level 2"
        .SeriesCollection(3).XValues = Sekt_Names()
        .SeriesCollection.NewSeries
        .SeriesCollection(4).Values = Series3
        .SeriesCollection(4).Name = "Spont Level 3"
        .SeriesCollection(4).XValues = Sekt_Names()
        .SetElement (msoElementLegendBottom)
        '.Legend.LegendEntries(1).Delete
        
        For i = .SeriesCollection.Count To 2 Step -1
            .SeriesCollection(i).ChartType = xlLineMarkers
            .SeriesCollection(i).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(i).Format.Fill.Visible = msoTrue
            '.SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
            '.SeriesCollection(i).Name = "Level " & i - 1
            '.SeriesCollection(i).MarkerSize = 10
            LogicTest = 0
            
        Vals = .SeriesCollection(i).Values
            For j = 1 To .SeriesCollection(i).Points.Count
                If Not Vals(j) = vbNullString Then
                    LogicTest = 1
                End If
            Next j
        
            If LogicTest = 0 Then
                .Legend.LegendEntries(i).Delete
            End If
        Next i
         .Legend.LegendEntries(1).Delete
    End With
      
iTop = iTop + 306 '297

Next sRow

Application.PrintCommunication = False
With WsChrt.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 0
End With
Application.PrintCommunication = True

'Sheets("Spont").Activate
Call ChangePointColor_DataCharts_all


End Sub

Sub Chart_for_date_one()

Dim nRow As Integer, nCol As Integer, i As Integer, sRow As Integer, txt As String, j As Integer
Dim Ws As Worksheet, Vals() As Variant, LogicTest As Integer
Dim sDate() As Integer, sNum() As Integer
Dim Sekt_Names() As String, Data_Series() As Variant, Series1() As Variant, Series2() As Variant, Series3() As Variant
Dim SheetExist As Boolean, YesNo As Integer
    
Set Ws = ActiveWorkbook.Sheets("Spont")

SheetExist = Charts.sheetExists(Charts_UserForm.ComboBox_Date.text)

If SheetExist = True Then
    YesNo = MsgBox("Chart for Date: " & Charts_UserForm.ComboBox_Date.text & " already exist" & vbNewLine & "To proceed this chart has to be deleted." & vbNewLine & "Delete chart?", vbExclamation + vbYesNo)
    If YesNo = vbYes Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ActiveWorkbook.Sheets(Charts_UserForm.ComboBox_Date.text).Delete
        Application.DisplayAlerts = True
    Else
        Call MsgBox("Existing chart was keept." & vbNewLine & "Closing macro.", vbInformation + vbOKOnly)
        Exit Sub
    End If
End If

'Charts_UserForm.ComboBox_Date.Value = "99-01-02" '<<<<<<<<<<<<<<<<<<< remember to change this

'sDate = Charts.First_cell_with_data(Ws)
'sNum = Charts.First_cell_with_number(Ws)

nCol = Ws.Cells(3, Ws.Columns.Count).End(xlToLeft).Column 'finds the last non empty cell in 1st column
nRow = Ws.Cells(Ws.Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 1st row

'Call Create_Alarm_Levels(nCol)

For i = 4 To nRow
    If Ws.Cells(i, 1).text = Charts_UserForm.ComboBox_Date.Value Then
        sRow = i '<<<<<<<<<<<<<<<<<<<<< sRow = search row - row which contains data for date for which chart is created
    End If
Next

Sekt_Names = Charts.Create_Sektion_Names(nCol, sRow)
Data_Series = Charts.Create_Data_For_Series(Sekt_Names, sRow, nCol, nRow, 3)

ReDim Series1(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series1(i) = Data_Series(0, i)
Next i

ReDim Series2(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series2(i) = Data_Series(1, i)
Next i

ReDim Series3(UBound(Data_Series, 2))
For i = 0 To UBound(Data_Series, 2)
    Series3(i) = Data_Series(2, i)
Next i
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>finished here at friday. problem with chart for date if we want all levels at the same chart not all profiles have the same nuber of levels<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count), Type:=xlChart
ActiveChart.Name = Charts_UserForm.ComboBox_Date.Value

    With ActiveChart
        .ChartType = xlLine
        .ChartStyle = 232
        .PlotArea.Format.Fill.Visible = msoFalse
        '.SetSourceData Source:=Ws.Range(Ws.Cells(sRow, 2), Ws.Cells(sRow, nCol))
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = Series1
        .SeriesCollection(1).XValues = Sekt_Names() 'Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, nCol))
        .HasTitle = True
        .ChartTitle.text = "Spont movement chart" & " - " & "Measurent Date:" & " " & Charts_UserForm.ComboBox_Date.Value
        .Axes(xlCategory).TickLabelPosition = xlLow
        .SeriesCollection(1).Values = Series2
        '.SeriesCollection(1).Name = "Spont Level 1"
        .SeriesCollection(1).ChartType = xlLineMarkers
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).Format.Fill.Visible = msoTrue
        .SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values = Series1
        .SeriesCollection(2).Name = "Spont Level 1"
        .SeriesCollection(2).XValues = Sekt_Names()
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Values = Series2
        .SeriesCollection(3).Name = "Spont Level 2"
        .SeriesCollection(3).XValues = Sekt_Names()
        .SeriesCollection.NewSeries
        .SeriesCollection(4).Values = Series3
        .SeriesCollection(4).Name = "Spont Level 3"
        .SeriesCollection(4).XValues = Sekt_Names()
        '.Legend.LegendEntries(1).Delete
    
    
        For i = .SeriesCollection.Count To 2 Step -1
            .SeriesCollection(i).ChartType = xlLineMarkers
            .SeriesCollection(i).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(i).Format.Fill.Visible = msoTrue
            '.SeriesCollection(1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
            '.SeriesCollection(i).Name = "Level " & i - 1
            '.SeriesCollection(i).MarkerSize = 10
            LogicTest = 0
            
        Vals = .SeriesCollection(i).Values
            For j = 1 To .SeriesCollection(i).Points.Count
                If Not Vals(j) = vbNullString Then
                    LogicTest = 1
                End If
            Next j
        
            If LogicTest = 0 Then
                .Legend.LegendEntries(i).Delete
            End If
        Next i
         .Legend.LegendEntries(1).Delete
    End With
Call ChangePointColor_DataCharts
'    For i = 2 To 5
'    If i <= 3 Then
'        With ActiveChart
'            .SeriesCollection.NewSeries
'            .SeriesCollection(i).XValues = Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, nCol))
'            .SeriesCollection(i).Values = Sheets("AlarmLevels").Range(Sheets("AlarmLevels").Cells(i + 9, 1), Sheets("AlarmLevels").Cells(i + 9, nCol - 1))
'            .SeriesCollection(i).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent3
'            .SeriesCollection(i).Format.Line.DashStyle = msoLineLongDash
'            .SeriesCollection(i).Name = "Alarm Level 1"
'            .SeriesCollection(i).ChartType = xlLine
'        End With
'    Else
'        With ActiveChart
'            .SeriesCollection.NewSeries
'            .SeriesCollection(i).XValues = Ws.Range(Ws.Cells(1, 2), Ws.Cells(1, nCol))
'            .SeriesCollection(i).Values = Sheets("AlarmLevels").Range(Sheets("AlarmLevels").Cells(i + 9, 1), Sheets("AlarmLevels").Cells(i + 9, nCol - 1))
'            .SeriesCollection(i).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
'            .SeriesCollection(i).Format.Line.DashStyle = msoLineLongDash
'            .SeriesCollection(i).Name = "Alarm Level 2"
'            .SeriesCollection(i).ChartType = xlLine
'        End With
'    End If
'Next
    
'    With ActiveChart
'        .Legend.LegendEntries(3).Delete
'        .Legend.LegendEntries(3).Delete
'    End With
               
End Sub
Function Create_Data_For_Series(Sekt_Names() As String, Date_Row As Integer, nCol As Integer, nRow As Integer, Lvl_Number As Integer) As Variant()

'>>>>>>>>>>>>>>>>>>> This function populates Matrix Results(). Results() contains data for series. Each row in matrix corresponds to 1 series (1 spot level - Every sektion the same lvl) _
                     If on some level in some sektion point does not exist the blank place is left there. Matrix Results is populeted accordingly to points names.

Dim i As Integer, j As Integer, k As Integer, Points_Row As Integer
Dim Ws As Worksheet
Dim Results() As Variant
ReDim Results(0 To Lvl_Number - 1, 0 To UBound(Sekt_Names))

Set Ws = Sheets("Spont")

'Points_Row = sNum(0) - 1

For i = 0 To UBound(Sekt_Names)
    For j = 0 To Lvl_Number - 1
        For k = 2 To nCol
            If Sekt_Names(i) & "." & j + 1 = Ws.Cells(3, k) Then
            Results(j, i) = Ws.Cells(Date_Row, k)
            End If
        Next k
    Next j
Next i

Create_Data_For_Series = Results()

End Function

Sub test()
Dim Point As String
Dim myXVal() As Double

Point = Cells(3, 2).text
myXVal = Charts.Alarm_Levels(Point)

'Call Charts.Alarm_Levels(Point)

Debug.Print myXVal(0)
Debug.Print myXVal(1)

End Sub
Function Alarm_Levels(Point_Name As String) As Double()
Dim SplitTxt() As String, Results(1) As Double

SplitTxt = Split(Point_Name, ".")

If SplitTxt(0) = 1 And SplitTxt(2) = 1 Or SplitTxt(0) = 2 And SplitTxt(2) = 1 Then 'Spontline 1 or 2 lvl 1
Results(0) = 0.04
Results(1) = 0.05
End If

If SplitTxt(0) = 1 And SplitTxt(2) = 2 Or SplitTxt(0) = 2 And SplitTxt(2) = 2 Then 'Spontline 1 or 2 lvl 2
Results(0) = 0.03
Results(1) = 0.045
End If

If SplitTxt(0) = 1 And SplitTxt(2) = 3 Or SplitTxt(0) = 2 And SplitTxt(2) = 3 Then 'Spontline 1 or 2 lvl 3
Results(0) = 0.02
Results(1) = 0.03
End If

If SplitTxt(0) = 3 Then 'Spontline 3 to 6 all lvl
Results(0) = 0.015
Results(1) = 0.02
End If

Alarm_Levels = Results()

End Function





Sub Create_Alarm_Levels(nCol As Integer)

Dim CtrlVal As Boolean

CtrlVal = Charts.WorkSheetExist("AlarmLevels")

If CtrlVal = False Then
    Sheets.Add.Name = "AlarmLevels"
End If

    Sheets("AlarmLevels").Cells(1, 1).Value = 0
    Sheets("AlarmLevels").Cells(2, 1).Value = 1
    
    Sheets("AlarmLevels").Cells(1, 2).Value = 0.02
    Sheets("AlarmLevels").Cells(2, 2).Value = 0.02
    Sheets("AlarmLevels").Cells(3, 2).Value = 0.03
    Sheets("AlarmLevels").Cells(4, 2).Value = 0.03
    Sheets("AlarmLevels").Cells(5, 2).Value = -0.02
    Sheets("AlarmLevels").Cells(6, 2).Value = -0.02
    Sheets("AlarmLevels").Cells(7, 2).Value = -0.03
    Sheets("AlarmLevels").Cells(8, 2).Value = -0.03
    
    Sheets("AlarmLevels").Cells(1, 3).Value = 0.03
    Sheets("AlarmLevels").Cells(2, 3).Value = 0.03
    Sheets("AlarmLevels").Cells(3, 3).Value = 0.045
    Sheets("AlarmLevels").Cells(4, 3).Value = 0.045
    Sheets("AlarmLevels").Cells(5, 3).Value = -0.03
    Sheets("AlarmLevels").Cells(6, 3).Value = -0.03
    Sheets("AlarmLevels").Cells(7, 3).Value = -0.045
    Sheets("AlarmLevels").Cells(8, 3).Value = -0.045
    
    Sheets("AlarmLevels").Cells(1, 4).Value = 0.04
    Sheets("AlarmLevels").Cells(2, 4).Value = 0.04
    Sheets("AlarmLevels").Cells(3, 4).Value = 0.05
    Sheets("AlarmLevels").Cells(4, 4).Value = 0.05
    Sheets("AlarmLevels").Cells(5, 4).Value = -0.04
    Sheets("AlarmLevels").Cells(6, 4).Value = -0.04
    Sheets("AlarmLevels").Cells(7, 4).Value = -0.05
    Sheets("AlarmLevels").Cells(8, 4).Value = -0.05
    
    Sheets("AlarmLevels").Cells(1, 5).Value = 0.015
    Sheets("AlarmLevels").Cells(2, 5).Value = 0.015
    Sheets("AlarmLevels").Cells(3, 5).Value = 0.02
    Sheets("AlarmLevels").Cells(4, 5).Value = 0.02
    Sheets("AlarmLevels").Cells(5, 5).Value = -0.015
    Sheets("AlarmLevels").Cells(6, 5).Value = -0.015
    Sheets("AlarmLevels").Cells(7, 5).Value = -0.02
    Sheets("AlarmLevels").Cells(8, 5).Value = -0.02
    

Sheets("Spont").Range(Sheets("Spont").Cells(1, 2), Sheets("Spont").Cells(1, nCol)).Copy Destination:=Sheets("AlarmLevels").Rows(10)


Sheets("AlarmLevels").Range(Sheets("AlarmLevels").Cells(11, 1), Sheets("AlarmLevels").Cells(11, nCol - 1)).Value = 0.02
Sheets("AlarmLevels").Range(Sheets("AlarmLevels").Cells(12, 1), Sheets("AlarmLevels").Cells(12, nCol - 1)).Value = -0.02
Sheets("AlarmLevels").Range(Sheets("AlarmLevels").Cells(13, 1), Sheets("AlarmLevels").Cells(13, nCol - 1)).Value = 0.03
Sheets("AlarmLevels").Range(Sheets("AlarmLevels").Cells(14, 1), Sheets("AlarmLevels").Cells(14, nCol - 1)).Value = -0.03

    Sheets("AlarmLevels").Visible = False
    
End Sub

Function WorkSheetExist(n As String) As Boolean
Dim Ws As Worksheet
  WorkSheetExist = False
  For Each Ws In Worksheets
    If n = Ws.Name Then
      WorkSheetExist = True
      Exit Function
    End If
  Next Ws
End Function

Sub First_cell_with_data()

Dim i As Integer, nRow As Integer, nCol As Integer, j As Integer
Dim Results(1) As Integer, Ws As Worksheet

Set Ws = Sheets("Spont")

For i = 1 To Ws.Columns.Count 'we loop throuh every column in the sheet
    nRow = Ws.Cells(Rows.Count, i).End(xlUp).Row 'for every column we find the last row which is not empty
    If nRow > 0 Then 'if column is not empty we go through every cell i that collumn (till last non empty row) and check if it contains date
        For j = 1 To nRow
            If IsDate(Cells(i, j)) = True Then
                If DatePart("yyyy", Cells(i, j)) > 1900 Then
                    Results(0) = i
                    Results(1) = j
                    'First_cell_with_data = Results()
                    Exit Sub      'at the first hit we end funcion
                End If
            End If
        Next j
    End If
Next i

End Sub

Sub First_cell_with_number()

Dim i As Integer, nRow As Integer, nCol As Integer, j As Integer
Dim Results(1) As Integer, Ws As Worksheet

Set Ws = Sheets("Spont")

For i = 1 To Ws.Columns.Count 'we loop throuh every column in the sheet
    nRow = Ws.Cells(Rows.Count, i).End(xlUp).Row 'for every column we find the last row which is not empty
    If nRow > 0 Then 'if column is not empty we go through every cell i that collumn (till last non empty row) and check if it contains date
        For j = 1 To nRow
            If Application.WorksheetFunction.IsNumber(Cells(i, j)) = True And Not (Cells(i, j)) = "" Then
                Results(0) = i
                Results(1) = j
                'First_cell_with_number = Results()
                Exit Sub      'at the first hit we end funcion
            End If
        Next j
    End If
Next i



End Sub
Function Create_Sektion_Names(nCol As Integer, sRow As Integer) As String()

Dim Results() As String, Arr() As String, ArrNext() As String
Dim i As Integer, j As Integer, k As Integer, n As Integer, StartingColumn As Integer
Dim Ws As Worksheet
Dim PtName As String
Dim Sekt() As String, Spont() As Integer, SpontNo As Integer, SektNo As Integer

Set Ws = Sheets("Spont")

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> creat matrix with empty places if level in profile is missing <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'If Charts_UserForm.ComboBox_Date.Value = "" Then
    StartingColumn = First_data_in_row(Ws.Cells(sRow, 1).Value)
'Else
    'StartingColumn = First_data_in_row(Charts_UserForm.ComboBox_Date.Value)
'End If

For i = StartingColumn To nCol - 1
    Arr = Split(Ws.Cells(3, i), ".")
    ArrNext = Split(Ws.Cells(3, i + 1), ".")
        If Arr(0) <> ArrNext(0) Then
           SpontNo = SpontNo + 1 '>>>>>makro counts how many times spont line number  changes. _
    >>>>>>>>>>>>>>>>>>>>>>>>>>>>For example 1.1 1.2 1.3 - the name changes 2 times from 1.1 to 1.2 and from 1.2 to 1.3, so if we want a nuber of section we must + 1
        End If
Next i
'SpontNo = SpontNo + 1 'now we know how many spont lines we have so we can resize Spont names Array
ReDim Spont(SpontNo)

For i = StartingColumn To nCol - 1
    Arr = Split(Ws.Cells(3, i), ".")
    ArrNext = Split(Ws.Cells(3, i + 1), ".")
        If Arr(0) <> ArrNext(0) Then
           Spont(j) = Arr(0) 'we repeat the same procedure as above but this time when first part of sektion name changes we save the spont line number _
                            ex: previous point is 1.4.3 next is 2.1.1 so Arr(0)=1 and ArrNext(0)=2 we save into Spont(0) = 1
           j = j + 1
        End If
Next i

If j = 0 Then
    Spont(j) = Arr(0)
Else
    Spont(j) = Spont(j - 1) + 1 'here we save last spont line name. For example in exel file we chave points only in spont lines 2 and 3, so the name change will occur only once from 2 to 3. _
                            Because of that we will save in Spont() on position Spont(0) number 2 and we need also 3 on position Spont(1)
End If

For j = 0 To UBound(Spont)
    For i = StartingColumn To nCol - 1
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
    For i = StartingColumn To nCol - 1
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




'ReDim Preserve Sekt(SpontNo)

'For j = 1 To SpontNo
    'For i = StartingColumn To nCol - 1
        'Arr = Split(Ws.Cells(3, i), ".")
    'ArrNext = Split(Ws.Cells(3, i + 1), ".")
        'If Arr(0) = j And ArrNext(0) = j And Arr(1) <> ArrNext(1) Then
        'Sekt(j) = Sekt(j) + 1
        'End If
    'Next i
    'Sekt(j) = Sekt(j) + 1 '>>>>>makro counts how many times sektion name changes. _
    >>>>>>>>>>>>>>>>>>>>>>>>>>>>For example 1.1 1.2 1.3 - the name changes 2 times from 1.1 to 1.2 and from 1.2 to 1.3, so if we want a nuber of section we must + 1
'Next j
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> crating matrix holding points number with extra points if in reality lvl 2 or 3 for some sektion doesn't exist _
need to do that so we have the same nuber of data with every series (levels in sektions makes series)<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'For i = 1 To SpontNo
    'SektNo = SektNo + Sekt(i)
'Next i

'ReDim Preserve Results(SektNo - 1)

'For i = 1 To SpontNo
    'On Error Resume Next
        'For k = 0 To UBound(Results)
            'If Results(k) = "" Then
            'n = k
            'Exit For
            'End If
        'Next k
    'For j = 1 To Sekt(i)
        'Results(n) = i & "." & j
        'n = n + 1
    'Next j
'Next i

Create_Sektion_Names = Sekt()

End Function
Function First_data_in_row(Datum As String) As Integer
Dim Ws As Worksheet, nRow As Integer, nCol As Integer, RowToCheck As Integer, i As Integer

Set Ws = Sheets("Spont")

nRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row 'finds the last non empty cell in 1st row
nCol = Ws.Cells(3, Columns.Count).End(xlToLeft).Column 'finds the last non empty cell in 1st column

For i = 4 To nRow
    If Ws.Cells(i, 1) = Datum Then
        RowToCheck = i
    End If
Next i

For i = 2 To nCol
    If IsNumeric(Ws.Cells(RowToCheck, i)) = True And IsEmpty(Ws.Cells(RowToCheck, i)) = False Then
        First_data_in_row = i
        Exit Function
    End If
Next i

End Function
Function Merge_Arrays(Arr1() As String, Arr2() As String, Rows As Integer) As String() 'don't know if this will be still necessary

Dim i As Integer, j As Integer
Dim Results() As String

If Arr1(1, 1) = 0 Then
    Merge_Arrays = Arr2()
    Else
    ReDim Preserve Arr1(Rows, UBound(Arr1) + UBound(Arr2))
    For i = UBound(Arr1) To UBound(Arr2)
        For j = 1 To Rows
        Arr1(j, i) = Arr2(j, i - UBound(Arr2))
        Next j
    Next i
    Merge_Arrays = Arr1()
End If

End Function

Sub ChangePointColor_DataCharts()
    Dim x As Integer, i As Integer, j As Integer
    Dim pointsNames As Variant, pointsValues As Variant
    Dim Ws As Worksheet
    Dim Point_Name As String
    Dim Larm_niva() As Double
    
    Set Ws = Worksheets("Spont")
    
    For i = 2 To ActiveChart.SeriesCollection.Count

    With ActiveChart.SeriesCollection(i)
        pointsNames = .XValues
        pointsValues = .Values
        For x = LBound(pointsNames) To UBound(pointsNames)
            Point_Name = pointsNames(x) & "." & i - 1
            Larm_niva = Alarm_Levels(Point_Name)
                If Abs(pointsValues(x)) < Larm_niva(0) Then
                    .Points(x).Format.Fill.ForeColor.RGB = RGB(0, 153, 0)
                    '.Points(x).Format.Fill.Solid
                Else
                    If Abs(pointsValues(x)) >= Larm_niva(1) Then
                        .Points(x).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                        '.Points(x).Format.Fill.Solid
                    Else
                        .Points(x).Format.Fill.ForeColor.RGB = RGB(255, 153, 51)
                        '.Points(x).Format.Fill.Solid
                    End If
                End If
        Next x
    End With
    
    Next i
    
End Sub
Sub ChangePointColor_DataCharts_all()
    Dim x As Integer, i As Integer, j As Integer
    Dim pointsNames As Variant, pointsValues As Variant
    Dim Ws As Worksheet, WsChrt As Worksheet
    Dim Point_Name As String
    Dim Larm_niva() As Double
    
    Set Ws = Worksheets("Spont")
    Set WsChrt = ActiveSheet
    
For j = 1 To WsChrt.ChartObjects.Count
    For i = 2 To WsChrt.ChartObjects(j).Chart.SeriesCollection.Count

    With WsChrt.ChartObjects(j).Chart.SeriesCollection(i)
        pointsNames = .XValues
        pointsValues = .Values
        For x = LBound(pointsNames) To UBound(pointsNames)
            Point_Name = pointsNames(x) & "." & i - 1
            Larm_niva = Alarm_Levels(Point_Name)
                If Abs(pointsValues(x)) < Larm_niva(0) Then
                    .Points(x).Format.Fill.ForeColor.RGB = RGB(0, 153, 0)
                    '.Points(x).Format.Fill.Solid
                Else
                    If Abs(pointsValues(x)) >= Larm_niva(1) Then
                        .Points(x).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                        '.Points(x).Format.Fill.Solid
                    Else
                        .Points(x).Format.Fill.ForeColor.RGB = RGB(255, 153, 51)
                        '.Points(x).Format.Fill.Solid
                    End If
                End If
        Next x
    End With
    
    Next i
Next j

End Sub
Sub ChangePointColor_DataCharts_Sections(sCol As Integer)
    Dim x As Integer, i As Integer, j As Integer
    Dim pointsValues() As Variant
    Dim Ws As Worksheet
    Dim Point_Name As String
    Dim Larm_niva() As Double
    Dim Split_sekt_Name() As String, Arr() As String
    
    Set Ws = Worksheets("Spont")
    
    For i = 1 To ActiveChart.SeriesCollection.Count
    
    With ActiveChart.SeriesCollection(i)
        'pointsNames = .XValues
        pointsValues = .Values
        For x = LBound(pointsValues) To UBound(pointsValues)
            Point_Name = Ws.Cells(3, sCol + i - 1)
            Larm_niva = Alarm_Levels(Point_Name)
                If Abs(pointsValues(x)) < Larm_niva(0) Then
                    .Points(x).Format.Fill.ForeColor.RGB = RGB(0, 153, 0)
                    '.Points(x).Format.Fill.Solid
                Else
                    If Abs(pointsValues(x)) >= Larm_niva(1) Then
                        .Points(x).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                        '.Points(x).Format.Fill.Solid
                    Else
                        .Points(x).Format.Fill.ForeColor.RGB = RGB(255, 153, 51)
                        '.Points(x).Format.Fill.Solid
                    End If
                End If
        Next x
    End With
    
    Next i
    
End Sub
Function sheetExists(sheetToFind As String) As Boolean
Dim Sheet As Worksheet, chSheet As Chart
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
    
    For Each chSheet In ActiveWorkbook.Charts
    If sheetToFind = chSheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next chSheet
    
End Function
