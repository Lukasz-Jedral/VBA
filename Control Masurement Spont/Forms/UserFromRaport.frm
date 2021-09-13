VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFromRaport 
   Caption         =   "Raport"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   OleObjectBlob   =   "UserFromRaport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFromRaport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OkButton_Click()

Unload Me

End Sub

Private Sub SaveButton_Click()
Dim File_Path As String

File_Path = Application.GetSaveAsFilename(meas_Date, "Text File (*.txt), *.txt")

Open File_Path For Output As #2
Write #2, msg
Close #2

Call MsgBox("Raport was saved", vbInformation + vbOKOnly)

Unload Me

End Sub

Private Sub UserForm_Initialize()


TextBox.Caption = msg
TextBox.AutoSize = True
UserFromRaport.Height = TextBox.Height + 50

End Sub
