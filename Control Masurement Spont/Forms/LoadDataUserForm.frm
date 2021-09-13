VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadDataUserForm 
   Caption         =   "Awsome Macro"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5325
   OleObjectBlob   =   "LoadDataUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadDataUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub ExitProgram_Click()
Me.Hide
End Sub


Private Sub LoadMeasurment_Click()

Call Displacments.Caculations

End Sub

Private Sub Ref_Line_Dir_Click()

End Sub

Private Sub ReferenceLineDirectory_Click()

Call Displacments.SaveRefFileDirectory(1)

Call UserForm_Initialize

End Sub

Private Sub ReferencePointsDirectory_Click()

Call Displacments.SaveRefFileDirectory(2)

Call UserForm_Initialize
    
End Sub

Private Sub UserForm_Initialize()
Dim Short_File_Name() As String, LineFromFile As String

If Dir(ActiveWorkbook.Path & "\Excel_Macro_Data\RefLineDir.txt", vbDirectory) = vbNullString Then
    With LoadDataUserForm.Ref_Line_Dir
        .Caption = "Not loaded"
        .ForeColor = &HFF&
    End With
Else
    Open ActiveWorkbook.Path & "\Excel_Macro_Data\RefLineDir.txt" For Input As #1
    Line Input #1, LineFromFile
    Short_File_Name = Split(LineFromFile, "\")
    Close #1
        If Short_File_Name(0) = Chr(34) & Chr(34) Then
            With LoadDataUserForm.Ref_Line_Dir
            .Caption = "Not loaded"
            .ForeColor = &HFF&
            End With
        Else
            With LoadDataUserForm.Ref_Line_Dir
            .Caption = Left(Short_File_Name(UBound(Short_File_Name)), Len(Short_File_Name(UBound(Short_File_Name))) - 1)
            .ForeColor = &H80000012
            End With
        End If
End If

If Dir(ActiveWorkbook.Path & "\Excel_Macro_Data\RefPointsDir.txt", vbDirectory) = vbNullString Then
    With LoadDataUserForm.Ref_Point_Dir
        .Caption = "Not loaded"
        .ForeColor = &HFF&
    End With
Else
    Open ActiveWorkbook.Path & "\Excel_Macro_Data\RefPointsDir.txt" For Input As #1
    Line Input #1, LineFromFile
    Short_File_Name = Split(LineFromFile, "\")
    Close #1
        If Short_File_Name(0) = Chr(34) & Chr(34) Then
            With LoadDataUserForm.Ref_Point_Dir
            .Caption = "Not loaded"
            .ForeColor = &HFF&
            End With
        Else
            With LoadDataUserForm.Ref_Point_Dir
            .Caption = Left(Short_File_Name(UBound(Short_File_Name)), Len(Short_File_Name(UBound(Short_File_Name))) - 1)
            .ForeColor = &H80000012
            End With
        End If
End If
 
End Sub


