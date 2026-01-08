VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPHOPDE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patroli Harian Operator (AC90)"
   ClientHeight    =   6690
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   11955
   ControlBox      =   0   'False
   Icon            =   "FrmPHOPDE.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "FrmPHOPDE.dsx":1084A
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmPHOPDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Private Sub CommandButton1_Click()
Dim TekananUtama
Dim TekananFB
Dim TekananCA
Dim TekananCB
Dim TekananRS
Dim KebersihanM
Dim PA5
Dim Paku
Dim Penggulung
Dim Lampu
Dim Hasil
Dim KonRollet
Dim Chiping
Dim Marker
Dim KonStraight
Dim Safety
Dim Ukuran
Dim TabungRS
Dim Equitment
Dim ccdd
Dim Datum

Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
If ComboBox1 = "" Then
MsgBox " NO MESIN belum terisi"
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
If OptionButton11.Value = False And OptionButton12.Value = False Then
MsgBox " No.7 belum terisi"
Exit Sub
End If
If OptionButton13.Value = False And OptionButton14.Value = False Then
MsgBox " No.8 belum terisi"
Exit Sub
End If
If OptionButton15.Value = False And OptionButton16.Value = False Then
MsgBox " No.9 belum terisi"
Exit Sub
End If
If OptionButton17.Value = False And OptionButton18.Value = False Then
MsgBox " No.10 belum terisi"
Exit Sub
End If
If OptionButton19.Value = False And OptionButton20.Value = False Then
MsgBox " No.11 belum terisi"
Exit Sub
End If
If OptionButton21.Value = False And OptionButton22.Value = False Then
MsgBox " No.12 belum terisi"
Exit Sub
End If
If OptionButton23.Value = False And OptionButton24.Value = False Then
MsgBox " No.13 belum terisi"
Exit Sub
End If
If OptionButton25.Value = False And OptionButton26.Value = False Then
MsgBox " No.14 belum terisi"
Exit Sub
End If
If OptionButton27.Value = False And OptionButton28.Value = False Then
MsgBox " No.15 belum terisi"
Exit Sub
End If
If OptionButton29.Value = False And OptionButton30.Value = False Then
MsgBox " No.16 belum terisi"
Exit Sub
End If
If OptionButton31.Value = False And OptionButton32.Value = False Then
MsgBox " No.17 belum terisi"
Exit Sub
End If
If OptionButton39.Value = False And OptionButton40.Value = False Then
MsgBox " No.18 belum terisi"
Exit Sub
End If
If OptionButton41.Value = False And OptionButton42.Value = False Then
MsgBox " No.19 belum terisi"
Exit Sub
End If
If OptionButton37.Value = False And OptionButton38.Value = False Then
MsgBox " No.6 belum terisi"
Exit Sub
End If
'If TextBox5 = "" Then
'MsgBox " No.21 belum terisi"
'Exit Sub
'End If
If OptionButton45.Value = False And OptionButton46.Value = False And OptionButton47.Value = False Then
MsgBox " No.21 belum terisi"
Exit Sub
End If

If OptionButton1.Value = True Then
TekananUtama = "O"
End If
If OptionButton2.Value = True Then
TekananUtama = "NG"
End If

If OptionButton3.Value = True Then
TekananFB = "O"
End If
If OptionButton4.Value = True Then
TekananFB = "NG"
End If

If OptionButton5.Value = True Then
TekananCA = "O"
End If
If OptionButton6.Value = True Then
TekananCA = "NG"
End If

If OptionButton7.Value = True Then
TekananCB = "O"
End If
If OptionButton8.Value = True Then
TekananCB = "NG"
End If

If OptionButton9.Value = True Then
TekananRS = "O"
End If
If OptionButton10.Value = True Then
TekananRS = "NG"
End If

If OptionButton37.Value = True Then
KebersihanM = "O"
End If
If OptionButton38.Value = True Then
KebersihanM = "NG"
End If

If OptionButton11.Value = True Then
PA5 = "O"
End If
If OptionButton12.Value = True Then
PA5 = "NG"
End If

If OptionButton13.Value = True Then
Paku = "O"
End If
If OptionButton14.Value = True Then
Paku = "NG"
End If

If OptionButton15.Value = True Then
Penggulung = "O"
End If
If OptionButton16.Value = True Then
Penggulung = "NG"
End If

If OptionButton17.Value = True Then
Lampu = "O"
End If
If OptionButton18.Value = True Then
Lampu = "NG"
End If

If OptionButton19.Value = True Then
Hasil = "O"
End If
If OptionButton20.Value = True Then
Hasil = "NG"
End If

If OptionButton21.Value = True Then
KonRollet = "O"
End If
If OptionButton22.Value = True Then
KonRollet = "NG"
End If

If OptionButton23.Value = True Then
Chiping = "O"
End If
If OptionButton24.Value = True Then
Chiping = "NG"
End If

If OptionButton25.Value = True Then
Marker = "O"
End If
If OptionButton26.Value = True Then
Marker = "NG"
End If

If OptionButton27.Value = True Then
KonStraight = "O"
End If
If OptionButton28.Value = True Then
KonStraight = "NG"
End If

If OptionButton29.Value = True Then
Safety = "O"
End If
If OptionButton30.Value = True Then
Safety = "NG"
End If

If OptionButton31.Value = True Then
Ukuran = "O"
End If
If OptionButton32.Value = True Then
Ukuran = "NG"
End If

If OptionButton39.Value = True Then
TabungRS = "O"
End If
If OptionButton40.Value = True Then
TabungRS = "NG"
End If

If OptionButton41.Value = True Then
Equitment = "O"
End If
If OptionButton42.Value = True Then
Equitment = "NG"
End If

If OptionButton43.Value = True Then
ccdd = "O"
End If
If OptionButton44.Value = True Then
ccdd = "NG"
End If

If OptionButton45.Value = True Then
Datum = "80 mm"
End If
If OptionButton46.Value = True Then
Datum = "100 mm"
End If
If OptionButton47.Value = True Then
Datum = "NG"
End If


If ComboBox3 = "" Then
MsgBox " SHIFT Belum terisi"
Exit Sub
End If

If TextBox3 = "" Then
MsgBox " TANGGAL belum terisi"
Exit Sub
End If

If TextBox4 = "" Then
MsgBox " NIK belum terisi"
Exit Sub
Else

kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbPHOP.mdb"
Dim SQLTambah As String
SQLTambah = "insert into tb1 (TANGGAL,NOMESIN,NIK,Shift,TekananUdaraUtama,TekananUdaraFeedBelt,TekananUdaraClampA,TekananUdaraClampB,TekananUdaraRubberSealUnit,KebersihanMesindanAplicator,PotOilPA5,PakuPengaman,PenggulungKertasTerminal,LampuAndonMesin,HasilCuttingStripping,KondisiRollet,ChipInsulationBoxTerminalBox,MarkershunkApplikator,KondisiStraightmeter,SafetyCover,UkuranCircuitYangTerproses,TabungRubberSeal,TidakAdaEquitmentYangMenghalangiOperator,CCDD,PanjangWireDatum,Keterangan) values ('" _
& TextBox3 & "','" & ComboBox1 & "','" & TextBox4 & "','" & ComboBox3 & "','" & TekananUtama & "','" & TekananFB & "','" & TekananCA & "','" & TekananCB & "','" & TekananRS & "','" & KebersihanM & "','" & PA5 & "','" & Paku & "','" & Penggulung & "','" & Lampu & "','" & Hasil & "','" & KonRollet & "','" & Chiping & "','" & Marker & "','" & KonStraight & "','" & Safety & "','" & Ukuran & "','" & TabungRS & "','" & Equitment & "','" & ccdd & "','" & Datum & "','" & TextBox2 & "')"


kon.Execute SQLTambah

'MsgBox "Input data berhasil", vbDefaultButton1

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

'TextBox3.Value = ""
ComboBox1.Value = ""
TextBox4.Value = ""
ComboBox3.Value = ""
OptionButton1.Value = "false"
OptionButton2.Value = "false"
OptionButton3.Value = "false"
OptionButton4.Value = "false"
OptionButton5.Value = "false"
OptionButton6.Value = "false"
OptionButton7.Value = "false"
OptionButton8.Value = "false"
OptionButton9.Value = "false"
OptionButton10.Value = "false"
OptionButton37.Value = "false"
OptionButton38.Value = "false"
OptionButton11.Value = "false"
OptionButton12.Value = "false"
OptionButton13.Value = "false"
OptionButton14.Value = "false"
OptionButton15.Value = "false"
OptionButton16.Value = "false"
OptionButton17.Value = "false"
OptionButton18.Value = "false"
OptionButton19.Value = "false"
OptionButton20.Value = "false"
OptionButton21.Value = "false"
OptionButton22.Value = "false"
OptionButton23.Value = "false"
OptionButton24.Value = "false"
OptionButton25.Value = "false"
OptionButton26.Value = "false"
OptionButton27.Value = "false"
OptionButton28.Value = "false"
OptionButton29.Value = "false"
OptionButton30.Value = "false"
OptionButton31.Value = "false"
OptionButton32.Value = "false"
OptionButton39.Value = "false"
OptionButton40.Value = "false"
OptionButton41.Value = "false"
OptionButton42.Value = "false"
TextBox2.Value = ""

Unload Me

End If
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton6_Click()
MsgBox "Data Berhasil Dihapus", vbInformation, "Informasi"
TextBox3.Value = ""
ComboBox1.Value = ""
TextBox4.Value = ""
ComboBox3.Value = ""
OptionButton1.Value = "false"
OptionButton2.Value = "false"
OptionButton3.Value = "false"
OptionButton4.Value = "false"
OptionButton5.Value = "false"
OptionButton6.Value = "false"
OptionButton7.Value = "false"
OptionButton8.Value = "false"
OptionButton9.Value = "false"
OptionButton10.Value = "false"
OptionButton37.Value = "false"
OptionButton38.Value = "false"
OptionButton11.Value = "false"
OptionButton12.Value = "false"
OptionButton13.Value = "false"
OptionButton14.Value = "false"
OptionButton15.Value = "false"
OptionButton16.Value = "false"
OptionButton17.Value = "false"
OptionButton18.Value = "false"
OptionButton19.Value = "false"
OptionButton20.Value = "false"
OptionButton21.Value = "false"
OptionButton22.Value = "false"
OptionButton23.Value = "false"
OptionButton24.Value = "false"
OptionButton25.Value = "false"
OptionButton26.Value = "false"
OptionButton27.Value = "false"
OptionButton28.Value = "false"
OptionButton29.Value = "false"
OptionButton30.Value = "false"
OptionButton31.Value = "false"
OptionButton32.Value = "false"
OptionButton39.Value = "false"
OptionButton40.Value = "false"
OptionButton41.Value = "false"
OptionButton42.Value = "false"
TextBox2.Value = ""
End Sub

Private Sub CommandButton9_Click()
DataPhOPOP.Show
End Sub

Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Text3.Text = Format(Date, "yyyy/mm/dd")
End Sub

Private Sub UserForm_Activate()
Module1.HideXCloseButton Me

Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\DbPHOP.mdb"

TextBox3.Text = Format(Now, "yyyy/mm/dd HH:MM")


Dim henti As Boolean
henti = False
Do Until henti
Label69.Caption = Now
DoEvents
Loop

End Sub
Private Sub UserForm_Initialize()

With ComboBox3
.AddItem "NS"
.AddItem "A"
.AddItem "B"
End With

With ComboBox1
.AddItem "AC90TRX28"
.AddItem "AC90TRX"
.AddItem "AC90NPR"
.AddItem "AC90FTZ"
.AddItem "AC90DFM"
.AddItem "AC90BIG"
.AddItem "AC90J72"
.AddItem "AC90J30"

End With

ComboBox1.Text = NoMesin.Text21.Text
End Sub
Private Sub ComboBox1_Change()
    ComboBox1.Text = UCase(ComboBox1.Text)
End Sub
Private Sub ComboBox3_Change()
    ComboBox3.Text = UCase(ComboBox3.Text)
End Sub

