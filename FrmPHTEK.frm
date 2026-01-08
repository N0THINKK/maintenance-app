VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPHTEK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patroli Harian (AC90)"
   ClientHeight    =   6840
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   11730
   ControlBox      =   0   'False
   Icon            =   "FrmPHTEK.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "FrmPHTEK.dsx":1084A
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmPHTEK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbPHTEK.mdb"
End Sub


Private Sub CommandButton1_Click()
Dim kon
Dim Nmr1
Dim Nmr2
Dim Nmr3
Dim Nmr4
Dim Nmr5
'Dim Nmr6
Dim Nmr7
Dim Nmr8
Dim Nmr9
Dim Nmr10
Dim Nmr11
Dim Nmr12
Dim Nmr13
Dim Nmr14
Dim Nmr15
Dim Nmr16
Dim Nmr17
Dim Nmr18
Dim Nmr19
Dim Nmr20

Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
If Text1 = "" Then
MsgBox "Tanggal belum diisi"
Exit Sub
End If

If ComboBox1 = "" Then
MsgBox "Nomer mesin belum terisi"
Exit Sub
End If

If ComboBox2 = "" Then
MsgBox "NIK belum terisi"
Exit Sub
End If

If Text2 = "" Then
MsgBox "Counter StrOe Blade belum diisi"
Exit Sub
End If

If OptionButton1.Value = False And OptionButton2.Value = False Then
MsgBox " No.1 belum terisi"
Exit Sub
End If
If OptionButton3.Value = False And OptionButton4.Value = False Then
MsgBox " No.2 belum terisi"
Exit Sub
End If
If OptionButton5.Value = False And OptionButton6.Value = False Then
MsgBox " No.3 belum terisi"
Exit Sub
End If
If OptionButton7.Value = False And OptionButton8.Value = False Then
MsgBox " No.4 belum terisi"
Exit Sub
End If
If OptionButton9.Value = False And OptionButton10.Value = False Then
MsgBox " No.5 belum terisi"
Exit Sub
End If
If OptionButton38.Value = False And OptionButton37.Value = False Then
MsgBox " No.7 belum terisi"
Exit Sub
End If
If OptionButton11.Value = False And OptionButton12.Value = False Then
MsgBox " No.8 belum terisi"
Exit Sub
End If
If OptionButton13.Value = False And OptionButton14.Value = False Then
MsgBox " No.9 belum terisi"
Exit Sub
End If
If OptionButton15.Value = False And OptionButton16.Value = False Then
MsgBox " No.10 belum terisi"
Exit Sub
End If
If OptionButton17.Value = False And OptionButton18.Value = False Then
MsgBox " No.11 belum terisi"
Exit Sub
End If
If OptionButton19.Value = False And OptionButton20.Value = False And OptionButton33.Value = False Then
MsgBox " No.12 belum terisi"
Exit Sub
End If
If OptionButton21.Value = False And OptionButton22.Value = False Then
MsgBox " No.13 belum terisi"
Exit Sub
End If
If OptionButton23.Value = False And OptionButton24.Value = False And OptionButton34.Value = False Then
MsgBox " No.14 belum terisi"
Exit Sub
End If
If OptionButton25.Value = False And OptionButton26.Value = False Then
MsgBox " No.15 belum terisi"
Exit Sub
End If
If OptionButton27.Value = False And OptionButton28.Value = False And OptionButton36.Value = False Then
MsgBox " No.16 belum terisi"
Exit Sub
End If
If OptionButton29.Value = False And OptionButton30.Value = False And OptionButton35.Value = False Then
MsgBox " No.17 belum terisi"
Exit Sub
End If
If OptionButton31.Value = False And OptionButton32.Value = False Then
MsgBox " No.18 belum terisi"
Exit Sub
End If
If OptionButton39.Value = False And OptionButton40.Value = False Then
MsgBox " No.19 belum terisi"
Exit Sub
End If
If OptionButton41.Value = False And OptionButton42.Value = False Then
MsgBox " No.20 belum terisi"
Exit Sub
End If


If OptionButton1.Value = True Then
Nmr1 = "O"
End If
If OptionButton2.Value = True Then
Nmr1 = "NG"
End If
If OptionButton3.Value = True Then
Nmr2 = "O"
End If
If OptionButton4.Value = True Then
Nmr2 = "NG"
End If
If OptionButton5.Value = True Then
Nmr3 = "O"
End If
If OptionButton6.Value = True Then
Nmr3 = "NG"
End If
If OptionButton7.Value = True Then
Nmr4 = "O"
End If
If OptionButton8.Value = True Then
Nmr4 = "NG"
End If
If OptionButton9.Value = True Then
Nmr5 = "O"
End If
If OptionButton10.Value = True Then
Nmr5 = "NG"
End If
If OptionButton38.Value = True Then
Nmr7 = "O"
End If
If OptionButton37.Value = True Then
Nmr7 = "NG"
End If

If OptionButton11.Value = True Then
Nmr8 = "O"
End If
If OptionButton12.Value = True Then
Nmr8 = "NG"
End If
If OptionButton13.Value = True Then
Nmr9 = "O"
End If
If OptionButton14.Value = True Then
Nmr9 = "NG"
End If
If OptionButton15.Value = True Then
Nmr10 = "O"
End If
If OptionButton16.Value = True Then
Nmr10 = "NG"
End If
If OptionButton17.Value = True Then
Nmr11 = "O"
End If
If OptionButton18.Value = True Then
Nmr11 = "NG"
End If
If OptionButton19.Value = True Then
Nmr12 = "O"
End If
If OptionButton20.Value = True Then
Nmr12 = "NG"
End If
If OptionButton33.Value = True Then
Nmr12 = "N/A"
End If
If OptionButton21.Value = True Then
Nmr13 = "O"
End If
If OptionButton22.Value = True Then
Nmr13 = "NG"
End If
If OptionButton23.Value = True Then
Nmr14 = "O"
End If
If OptionButton24.Value = True Then
Nmr14 = "NG"
End If
If OptionButton34.Value = True Then
Nmr14 = "N/A"
End If
If OptionButton25.Value = True Then
Nmr15 = "O"
End If
If OptionButton26.Value = True Then
Nmr15 = "NG"
End If
If OptionButton27.Value = True Then
Nmr16 = "O"
End If
If OptionButton28.Value = True Then
Nmr16 = "NG"
End If
If OptionButton36.Value = True Then
Nmr16 = "N/A"
End If
If OptionButton29.Value = True Then
Nmr17 = "O"
End If
If OptionButton30.Value = True Then
Nmr17 = "NG"
End If
If OptionButton35.Value = True Then
Nmr17 = "N/A"
End If
If OptionButton31.Value = True Then
Nmr18 = "O"
End If
If OptionButton32.Value = True Then
Nmr18 = "NG"
End If
If OptionButton39.Value = True Then
Nmr19 = "O"
End If
If OptionButton40.Value = True Then
Nmr19 = "NG"
End If
If OptionButton41.Value = True Then
Nmr20 = "O"
End If
If OptionButton42.Value = True Then
Nmr20 = "NG"
End If



kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbPHTEK.mdb"
Dim SQLTambah As String
SQLTambah = "insert into tb1 (TANGGAL,NOMESIN,NIK,SensorWireTerbelit,WireGuide,SensorJoint,GuidePipe,CorrectionRoller,CounterStrokeBlade,CounterStrokeAplikator,SensorStrippingSisiAdanB,MK30,DeteksiCFM,DC10,CCDD,HasilStripping,LEDSensorSeal,KebocoranAngin,KecepatanNaikTurunCrimperSisiA,KecepatanNaikTurunCrimperSisiB,SensorWireHabis,TidakAdaForeignMaterialPadaStopKontak,KebersihanStopKontakDanSteker,Keterangan) values ('" _
& Text1 & "','" & ComboBox1 & "','" & ComboBox2 & "','" & Nmr1 & "','" & Nmr2 & "','" & Nmr3 & "','" & Nmr4 & "','" & Nmr5 & "','" & Text2 & "','" & Nmr7 & "','" & Nmr8 & "','" & Nmr9 & "','" & Nmr10 & "','" & Nmr11 & "','" & Nmr12 & "','" & Nmr13 & "','" & Nmr14 & "','" & Nmr15 & "','" & Nmr16 & "','" & Nmr17 & "','" & Nmr18 & "','" & Nmr19 & "','" & Nmr20 & "','" & Text3 & "')"

kon.Execute SQLTambah

'MsgBox "Input data berhasil", vbDefaultButton1

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

'Text1 = ""
Text2 = ""
Text3 = ""
'ComboBox1.Value = ""  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc
'ComboBox2.Value = ""
'OptionButton1.Value = False
'OptionButton2.Value = False
'OptionButton3.Value = False
'OptionButton4.Value = False
'OptionButton5.Value = False
'OptionButton6.Value = False
'OptionButton7.Value = False
'OptionButton8.Value = False
'OptionButton9.Value = False
'OptionButton10.Value = False
'OptionButton11.Value = False
'OptionButton12.Value = False
'OptionButton13.Value = False
'OptionButton14.Value = False
'OptionButton15.Value = False
'OptionButton16.Value = False
'OptionButton17.Value = False
'OptionButton18.Value = False
'OptionButton19.Value = False
'OptionButton20.Value = False
'OptionButton21.Value = False
'OptionButton22.Value = False
'OptionButton23.Value = False
'OptionButton24.Value = False
'OptionButton25.Value = False
'OptionButton26.Value = False
'OptionButton27.Value = False
'OptionButton28.Value = False
'OptionButton29.Value = False
'OptionButton30.Value = False
'OptionButton31.Value = False
'OptionButton32.Value = False
'OptionButton33.Value = False
'OptionButton34.Value = False
'OptionButton35.Value = False
'OptionButton36.Value = False

'Unload Me

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton6_Click()
Dim hps As String
'hps = "delete from tb1 where TANGGAL = '" & Text1 & "'"
'kon.Execute hps
MsgBox "Data Berhasil di Hapus", vbInformation, "we ada informasi"
Text1 = ""
Text2 = ""
Text3 = ""
ComboBox1 = ""
ComboBox2 = ""
ComboBox1.Value = ""
ComboBox2.Value = ""
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False
OptionButton4.Value = False
OptionButton5.Value = False
OptionButton6.Value = False
OptionButton7.Value = False
OptionButton8.Value = False
OptionButton9.Value = False
OptionButton10.Value = False
OptionButton11.Value = False
OptionButton12.Value = False
OptionButton13.Value = False
OptionButton14.Value = False
OptionButton15.Value = False
OptionButton16.Value = False
OptionButton17.Value = False
OptionButton18.Value = False
OptionButton19.Value = False
OptionButton20.Value = False
OptionButton21.Value = False
OptionButton22.Value = False
OptionButton23.Value = False
OptionButton24.Value = False
OptionButton25.Value = False
OptionButton26.Value = False
OptionButton27.Value = False
OptionButton28.Value = False
OptionButton29.Value = False
OptionButton30.Value = False
OptionButton31.Value = False
OptionButton32.Value = False
OptionButton33.Value = False
OptionButton34.Value = False
OptionButton35.Value = False
OptionButton36.Value = False


Form_Activate
End Sub

Private Sub CommandButton7_Click()
DataPHTK.Show
End Sub

Private Sub Text1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Text1.Text = Format(Date, "yyyy/mm/dd")
End Sub

Private Sub UserForm_Activate()
Module1.HideXCloseButton Me

Text1.Text = Format(Date, "yyyy/mm/dd")
Dim henti As Boolean
henti = False
Do Until henti
Waktu.Caption = Now
DoEvents
Loop
End Sub

 Private Sub Form_Resize()
On Error GoTo err
If Me.Width >= (1400 * Screen.TwipsPerPixelX) Then
   Me.Width = (1400 * Screen.TwipsPerPixelX)
   End If
   If Me.Height >= (600 * Screen.TwipsPerPixelX) Then
      Me.Height = (600 * Screen.TwipsPerPixelX)
      End If
      Exit Sub
err:
    Me.WindowState = 0
End Sub

Private Sub ComboBox1_Change()
    ComboBox1.Text = UCase(ComboBox1.Text)
End Sub

Private Sub UserForm_Initialize()

ComboBox1.Text = NoMesin.Text21.Text

NIKTeknisi

End Sub


Public Sub NIKTeknisi()

Set con = New ADODB.Connection
Set rs2 = New ADODB.Recordset
con.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs2.Open "select*from NIKPATROL", con, adOpenDynamic, adLockOptimistic
ComboBox2.Clear
Do While Not rs2.EOF
ComboBox2.AddItem rs2!NIK
rs2.MoveNext
Loop
End Sub
