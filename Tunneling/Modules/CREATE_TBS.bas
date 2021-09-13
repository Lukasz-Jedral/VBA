Attribute VB_Name = "CREATE_TBS"
Option Explicit
Sub CREATE_TBS()

Dim New_File_Path As String, File_Path As String, Folder_Path As String, RoadLinePath As String, TodaysDate As String, Section As String
Dim TunFileList() As Variant, TbsFile() As Variant, TempArray() As Variant
Dim i As Integer, j As Integer
Dim fso As FileSystemObject
Dim fileStream As TextStream
Const myDecSep As String = "."

'create empty TBS - it will be our default folder
'Set fileStream = fso.CreateTextFile(NewFilePath)

TodaysDate = Format(Now, "YYYY-MM-DD")

MsgBox ("Don't worry about .xlsm extension. It will be changed to .tbs automaticly")
New_File_Path = INPUT_DATA.SaveFilePath

File_Path = Left(New_File_Path, InStrRev(New_File_Path, ".") - 1)

File_Path = File_Path & ".tbs"

Folder_Path = Left(File_Path, InStrRev(File_Path, "\"))

Call INPUT_DATA.CreateAfile(File_Path) ' creating empty .tbs file

RoadLinePath = INPUT_DATA.GetTheFilePathExtended("Choose RoadLine for Tunnel Calculations", "Road Lines", "*.l3d", Folder_Path)

TunFileList = INPUT_DATA.Get_Folder_Content(Folder_Path) 'getting file paths for tun files

'filling tbs file with content

ReDim Preserve TbsFile(8)
TbsFile(1) = "CALC_DESCRIPTION,V2.00," & TodaysDate & ","
TbsFile(2) = ","
TbsFile(3) = "PTH=" & Right(RoadLinePath, Len(RoadLinePath) - Len(Folder_Path))
TbsFile(4) = "LIN=" & Right(RoadLinePath, InStrRev(RoadLinePath, "\"))
TbsFile(5) = "OFS=0.0"
TbsFile(6) = "GFA=1.0;1.0;1.0;1.0;1.0"
TbsFile(7) = "UNH=;0"
TbsFile(8) = "TUN=0.0;;;"

j = UBound(TbsFile) + 1

ReDim Preserve TbsFile(9 + 2 * UBound(TunFileList))
For i = 1 To UBound(TunFileList)
    TbsFile(j) = "PTH=" & Right(TunFileList(i), Len(TunFileList(i)) - Len(Folder_Path))
    j = j + 1
    Section = Replace(Format(Mid(TunFileList(i), InStrRev(TunFileList(i), "_") + 1, Len(Right(TunFileList(i), Len(TunFileList(i)) - InStrRev(TunFileList(i), "_") - 4))), "#0.000"), Application.DecimalSeparator, myDecSep) 'little trick for displaing dot as decimal sign regardless regional settings
    TbsFile(j) = "SEC=" & Section & ";;10.0;sb_225160,146;;;1;TUN"
    j = j + 1
Next i

TbsFile(UBound(TbsFile)) = "PRI=R;S;U1;U2"

For i = 1 To UBound(TbsFile)
    
Next i

'saving tbs file

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.CreateTextFile(File_Path, True) 'this line creates file
        
        For i = 1 To UBound(TbsFile)
         fileStream.Write TbsFile(i) & vbNewLine
        
        Next i
    fileStream.Close

MsgBox ("New Tbs file was created")

End Sub

Sub ShowUserForm()

InitTBS.Show

End Sub


