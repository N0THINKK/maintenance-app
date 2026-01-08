VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmHistry10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Problem Mesin 10"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   10560
   ControlBox      =   0   'False
   Icon            =   "FrmHistry10.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "FrmHistry10.dsx":1084A
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmHistry10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public con As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset


Private Sub Form_Activate()
Set kon = New ADODB.Connection
Set con = New ADODB.Connection
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset

kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Histry.mdb"

End Sub


Private Sub CheckBox2_Click()

If CheckBox2 = True Then
TextBox2.Value = Format(Now, "hh.mm.ss")
TextBox3.Value = TextBox1
ComboBox6.Value = "  "
TextBox7.Value = "  "
TextBox11.Value = "  "
TextBox8.Value = ""
ComboBox5.Value = "Perbaikan Berlanjut shift berikutnya"
TextBox4.Value = Format(Now, "hh.mm.ss")
ComboBox7.Value = "  "
End If

End Sub

Private Sub CheckBox3_Click()
If CheckBox3 = True Then
TextBox2.Value = Format(Now, "hh.mm.ss")
TextBox3.Value = Format(Now, "hh.mm.ss")
ComboBox3.Value = "Preventive Machine"
ComboBox4.Value = "    "
TextBox7.Value = "  "
TextBox11.Value = "  "
TextBox8.Value = ""
ComboBox6.Value = "Preventif Machine"
TextBox4.Value = Format(Now, "hh.mm.ss")
ComboBox7.Value = "  "
TextBox1.Value = "09.09.09"

End If
End Sub

Private Sub CommandButton1_Click()

Dim Ganti

SelisihPerbaikan
SelisihMesin

If ComboBox4 = "" Then
MsgBox " Data NIK belum terisi"
Exit Sub
End If

If ComboBox2 = "" Then
MsgBox " NO.Mesin Belum terisi"
Exit Sub
End If

If ComboBox3 = "" Then
MsgBox " Problem mesin belum terisi"
Exit Sub
End If

If TextBox10 = "" Then
MsgBox " TANGGAL belum terisi"
Exit Sub
End If

If TextBox1 = "" Then
MsgBox " START Problem belum terisi"
Exit Sub
End If

If TextBox2 = "" Then
MsgBox " STOP Problem belum terisi"
Exit Sub
End If

If TextBox3 = "" Then
MsgBox " START Repair belum terisi"
Exit Sub
End If

If ComboBox6 = "" Then
MsgBox " PENYEBAB Problem belum terisi"
Exit Sub
End If

If TextBox4 = "" Then
MsgBox " END Repair belum terisi"
Exit Sub
End If

If ComboBox7 = "" Then
MsgBox " PIC belum terisi"
Exit Sub
End If
If ComboBox5 = "" Then
MsgBox " TINDAKAN Perbaikan belum terisi"
Exit Sub
End If

Dim Panjang As Integer
Panjang = Len(ComboBox5.Text)
If Panjang <= 8 Then
MsgBox "Mohon isi Dengan lengkap yaa mas... " & vbCrLf & "" & vbCrLf & "Sedikit demi sedikit, lama-lama menjadi bukit " & vbCrLf & "Tak lengkap isi History, akan susah ketika Audit "
Exit Sub
End If


If CheckBox1 = True Then
Ganti = "4M"
FrmUtama.CommandButton13.Visible = True
Else
Ganti = ""
End If

If ComboBox8 = "" Then
MsgBox " Jenis Problem belum terisi"
Exit Sub
End If

kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\Histry.mdb"
Dim SQLTambah As String
SQLTambah = "insert into HS (TANGGAL,NIK,NOMESIN,StopMesin,StartMesin,ProblemMesin,StartRepair,PenyebabProblem,WaktuAmbilPart,NomerAplikator,JenisProblem,PartYangDiganti,TindakanPerbaikan,CounterStroke,EndRepair,PICRepair,GantiPart,Bulan,Tahun,DowntimeMesin,WaktuRepair) values ('" _
& TextBox10 & "','" & ComboBox4 & "','" & ComboBox2 & "','" & TextBox1 & "','" & TextBox2 & "','" & ComboBox3 & "','" & TextBox3 & "','" & ComboBox6 & "','" & TextBox7 & "','" & TextBox11 & "','" & ComboBox8 & "','" & TextBox8 & "','" & ComboBox5 & "','" & TextBox14 & "','" & TextBox4 & "','" & ComboBox7 & "','" & Ganti & "','" & Label62 & "','" & Label63 & "','" & TextBox12 & "','" & TextBox13 & "')"

kon.Execute SQLTambah

  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Data berhasil disimpan", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Shell App.Path & "\SimpanData.bat", vbMinimizedFocus

Unload Me
End Sub


Private Sub CommandButton10_Click()
DataHistryOP.Show
End Sub


Sub Clear_Control()
Text1.Text = ""
Text2.Text = ""
End Sub



Private Sub CommandButton11_Click()

End Sub

Private Sub CommandButton12_Click()
Set rs2 = New ADODB.Recordset
rs2.Open "select*from HS", kon, adOpenDynamic, adLockOptimistic
If Label63 = DataHistry.DataGrid1.Columns(0) And Label62 = DataHistry.DataGrid1.Columns(1) And Text11 = DataHistry.DataGrid1.Columns(2) And Combo1 = DataHistry.DataGrid1.Columns(4) And Combo2 = DataHistry.DataGrid1.Columns(6) Then
DataHistry.DataGrid1.Columns(10) = Text14
DataHistry.DataGrid1.Columns(11) = Text128
DataHistry.DataGrid1.Columns(12) = Text15
DataHistry.DataGrid1.Columns(13) = Text129

   MsgBox "Data sudah diupdate", vbInformation, "Informasi"
Else
MsgBox "Input tidak sesuai", vbInformation, "Informasi"
End If
End Sub

Private Sub CommandButton13_Click()
rs3.CursorLocation = adUseClient
rs3.Open "select * from HS where Tahun = '" & Label37 & "' and Bulan = '" & Label38 & "' and Tanggal = '" & Text11 & "' and NoMesin = '" & Combo1 & "' and Shift = '" & Combo2 & "'", con
If Not rs3.EOF Then
    Text14 = rs3!Qtyjam
    Text128 = rs3!Targetjam
    Text15 = rs3!EfisiensiRatio
    Text129 = rs3!WaktuMulai
    Text130 = rs3!WaktuMonitor
    Text13 = rs3!ExcludingTime
    Text131 = rs3!QtyCutting
    Text118 = rs3!SupportMH
    Text119 = rs3!Side
    Text120 = rs3!Color
    Text121 = rs3!SisiA
    Text122 = rs3!SisiB
    Text123 = rs3!Dobel
    Text124 = rs3!TungguMaterial
    Text125 = rs3!MesinTrouble
    Text126 = rs3!TungguKanban
    Text127 = rs3!Other
    
MsgBox "Data tersimpan Berhasil di Tampilkan", vbInformation, "Informasi"
Else
    MsgBox "Data yang anda cari tidak ada", vbInformation, "Ada Informasi!!!!!"
End If
End Sub


Private Sub CommandButton2_Click()
MsgBox "Input data dihapus", vbDefaultButton1
ComboBox4.Value = ""
ComboBox2.Value = ""
'TextBox1.Value = ""
'TextBox2.Value = ""
ComboBox3.Value = ""
'TextBox3.Value = ""
ComboBox6.Value = ""
TextBox7.Value = ""
TextBox11.Value = ""
TextBox8.Value = ""
ComboBox5.Value = ""
'TextBox4.Value = ""
ComboBox7.Value = ""

End Sub

Private Sub CommandButton9_Click()
Dim isi As Boolean
If ComboBox4.Value = "" Then
isi = True
Else
isi = False
End If


If ComboBox3.Value = "" Then
isi = True
Else
isi = False
End If

If isi = True Then
Unload Me
Else
MsgBox " Input Belum Selesai "
End If
End Sub




Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox1 = Format(Now, "hh.mm.ss")
End Sub

Private Sub TextBox10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox10.Text = Format(Date, "yyyy/mm/dd")
End Sub

Private Sub TextBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox2.Text = Format(Now, "hh.mm.ss")
End Sub

Private Sub TextBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox3.Text = Format(Now, "hh.mm.ss")
End Sub

Private Sub TextBox4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox4.Text = Format(Now, "hh.mm.ss")
End Sub

Private Sub UserForm_Activate()
Module1.HideXCloseButton Me

TextBox1 = Format(Now, "hh.mm.ss")

'If TextBox10.Value = "" Then
TextBox10.Text = Format(Date, "yyyy/mm/dd")
'Else
'TextBox10.Text = no1
'End If


Label15.Caption = Format(Now, "yyyy/mm/dd hh.mm")
Label62.Caption = Format(Now, "mm")
Label63.Caption = Format(Now, "yyyy")

Dim Berhenti As Boolean
Berhenti = False
Do Until Berhenti
TextBox2.Text = Format(Now, "hh.mm.ss")
DoEvents
Loop

End Sub

Private Sub ComboBox2_Change()
    ComboBox2.Text = UCase(ComboBox2.Text)
End Sub
Private Sub ComboBox7_Change()
    ComboBox7.Text = UCase(ComboBox7.Text)
End Sub

Private Sub Form_Resize()

    Dim D(4) As Double
    Dim i As Long
    Dim TempPoz As Long
    Dim StartPoz As Long
    Dim Ctl As Control
    Dim TempVisible As Boolean
    Dim ScaleX As Double
    Dim ScaleY As Double
    'Hitung skala-nya
    ScaleX = ScaleWidth / lngFormWidth
    ScaleY = ScaleHeight / lngFormHeight
    On Error Resume Next
    'Untuk setiap control yang terdapat di form
    For Each Ctl In Me
        TempVisible = Ctl.Visible
        Ctl.Visible = False
        StartPoz = 1
        'Baca data dari property Tag
        For i = 0 To 4
            TempPoz = InStr(StartPoz, Ctl.Tag, " ", _
                vbTextCompare)
            If TempPoz > 0 Then
                D(i) = Mid(Ctl.Tag, StartPoz, _
                    TempPoz - StartPoz)
                StartPoz = TempPoz + 1
            Else
                D(i) = 0
            End If
            'Pindahkan control berdasarkan data
            'di property Tag dan di skala form
            Ctl.Move D(0) * ScaleX, D(1) * ScaleY, _
                D(2) * ScaleX, D(3) * ScaleY
            Ctl.Width = D(2) * ScaleX
            Ctl.Height = D(3) * ScaleY
            'Ganti ukuran huruf
            If ScaleX < ScaleY Then
                   Ctl.FontSize = D(4) * ScaleX
            Else
                   Ctl.FontSize = D(4) * ScaleY
            End If
        Next i
        Ctl.Visible = TempVisible
    Next Ctl
    On Error GoTo 0
End Sub


Private Sub UserForm_Initialize()
Module1.HideXCloseButton Me

With ComboBox2
.AddItem "AC90TRX28"
.AddItem "AC90TRX"
.AddItem "AC90NPR"
.AddItem "AC90FTZ"
.AddItem "AC90DFM"
.AddItem "AC90BIG"
.AddItem "AC90J72"
.AddItem "AC90J30"
End With

With ComboBox8
.AddItem "Aplikator"
.AddItem "Servo"
.AddItem "Cutting / Stripping NG"
.AddItem "Rubber Seal"
.AddItem "CPU / Monitor problem"
.AddItem "CFM error"
.AddItem "Other"
End With

ComboBox2.Text = NoMesin.Text21.Text

ProblemMesin
PenyebabProblem
Perbaikan
PICRepair
NIKOP


End Sub


Public Sub ProblemMesin()

Set don = New ADODB.Connection
Set rs3 = New ADODB.Recordset
don.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs3.Open "select*from Problem", don, adOpenDynamic, adLockOptimistic
ComboBox3.Clear
Do While Not rs3.EOF
ComboBox3.AddItem rs3!PROBLEM
rs3.MoveNext

Loop
End Sub

Public Sub PenyebabProblem()

Set don = New ADODB.Connection
Set rs3 = New ADODB.Recordset
don.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs3.Open "select*from Sebab", don, adOpenDynamic, adLockOptimistic
ComboBox6.Clear
Do While Not rs3.EOF
ComboBox6.AddItem rs3!Penyebab
rs3.MoveNext

Loop
End Sub

Public Sub Perbaikan()

Set don = New ADODB.Connection
Set rs3 = New ADODB.Recordset
don.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs3.Open "select*from Cara", don, adOpenDynamic, adLockOptimistic
ComboBox5.Clear
Do While Not rs3.EOF
ComboBox5.AddItem rs3!Tindakan
rs3.MoveNext

Loop
End Sub

Public Sub PICRepair()

Set don = New ADODB.Connection
Set rs3 = New ADODB.Recordset
don.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs3.Open "select*from PICREPAIR", don, adOpenDynamic, adLockOptimistic
ComboBox7.Clear
Do While Not rs3.EOF
ComboBox7.AddItem rs3!TEKNISI
rs3.MoveNext

Loop
End Sub

Public Sub NIKOP()

Set don = New ADODB.Connection
Set rs3 = New ADODB.Recordset
don.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\ListHistry.mdb"

rs3.Open "select*from NIKOP", don, adOpenDynamic, adLockOptimistic
ComboBox4.Clear
Do While Not rs3.EOF
ComboBox4.AddItem rs3!OPERATOR
rs3.MoveNext

Loop
End Sub

Public Sub SelisihMesin()
Dim mesin

If TextBox1.Text = "" Or TextBox2.Text = "" Then
SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Start/Stop Problem tidak boleh kosong", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else
mesin = CDate(CDate(TextBox2) - CDate(TextBox1))
TextBox12 = Format(mesin, "hh.mm.ss")
End If

End Sub

Public Sub SelisihPerbaikan()
Dim Perbaikan

If TextBox3.Text = "" Or TextBox4.Text = "" Then
SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Start/Stop Repair tidak boleh kosong", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else

Perbaikan = CDate(CDate(TextBox4) - CDate(TextBox3))
TextBox13 = Format(Perbaikan, "hh.mm.ss")
End If
End Sub
