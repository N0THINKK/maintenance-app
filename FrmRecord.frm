VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Record"
   ClientHeight    =   5445
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   8340
   ControlBox      =   0   'False
   Icon            =   "FrmRecord.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "FrmRecord.dsx":1084A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
DataPhOPGL.Show
End Sub

Private Sub CommandButton10_Click()
DataMikrometer.Show
End Sub

Private Sub CommandButton11_Click()
DataDefect.Show
End Sub

Private Sub CommandButton12_Click()
DataDailyReport.Show
End Sub

Private Sub CommandButton2_Click()
FrmUtama.Show
Unload Me
End Sub

Private Sub CommandButton6_Click()
DataHistry.Show
End Sub

Private Sub CommandButton7_Click()
DataPHTK.Show
End Sub

Private Sub CommandButton8_Click()
DataLKO.Show
End Sub

Private Sub CommandButton9_Click()
DataPHApplA.Show

End Sub

Private Sub UserForm_Activate()
Module1.HideXCloseButton Me
End Sub

