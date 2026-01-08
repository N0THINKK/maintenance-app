VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   5730
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   6525
   ControlBox      =   0   'False
   Icon            =   "Frm_login.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "Frm_login.dsx":1084A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kon As New ADODB.Connection
Public rs2 As New ADODB.Recordset

Dim keluar As Byte
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
TextBox2.PasswordChar = ""
Else
TextBox2.PasswordChar = "*"
End If
End Sub

Private Sub ComboBox1_Change()
Label12.Caption = ComboBox1
End Sub

Private Sub CommandButton1_Click()

kriteria

If ComboBox1.Text = "" Or ComboBox1.Text = "Operator" And TextBox2.Text = "" Then

FrmUtama.CommandButton10.Enabled = False
FrmUtama.CommandButton11.Enabled = False
FrmUtama.CommandButton12.Enabled = False
DataLKO.DataGrid1.AllowUpdate = False
DataLKOOP.DataGrid1.AllowUpdate = False
DataPHApplA.DataGrid1.AllowUpdate = False
DataPHApplB.DataGrid1.AllowUpdate = False
DataPhOPOP.DataGrid1.AllowUpdate = False
DataPhOPGL.DataGrid1.AllowUpdate = False
FrmPHOPA.TextBox3.Enabled = False
FrmPHOPB.TextBox3.Enabled = False
FrmPHOPDE.TextBox3.Enabled = False
FrmHistry1.TextBox10.Enabled = False
FrmUtama.Show
Unload Me
Else

If ComboBox1.Text = "GL Produksi" And TextBox2.Text = Paswd.Text6 Then

FrmUtama.CommandButton11.Enabled = False
DataLKO.DataGrid1.AllowUpdate = False
DataLKOOP.DataGrid1.AllowUpdate = False
Unload Me
FrmUtama.Show

Else

If ComboBox1.Text = "Teknisi" And TextBox2.Text = Paswd.Text7 Then

FrmUtama.CommandButton11.Enabled = False
DataLKO.DataGrid1.AllowUpdate = False
DataLKOOP.DataGrid1.AllowUpdate = False
Unload Me
FrmUtama.Show

Else

If ComboBox1.Text = "Administrator" And TextBox2.Text = Paswd.Text8 Then

Unload Me
FrmUtama.Show


Else

SetTimer hwnd, NV_CLOSEMSGBOX, 1000&, AddressOf TimerProc

  Call MessageBox(hwnd, "Password Salah ", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

TextBox2 = ""
End If
End If
End If
End If

End Sub

Private Sub CommandButton2_Click()
End
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

Private Sub CommandButton3_Click()
Kon2.Show
End Sub

Private Sub OptionButton1_Click()
If OptionButton1.Value = True Then
Label13.Caption = "Pagi"
End If

End Sub

Private Sub OptionButton2_Click()

If OptionButton2.Value = True Then
Label13.Caption = "Malam"
End If

End Sub

Private Sub OptionButton3_Click()
If OptionButton3.Value = True Then
Label13.Caption = "NS"
End If
End Sub

Private Sub UserForm_Activate()

Label8.Caption = Kon2.DataGrid3.Columns(0)
TextBox3.Text = Kon2.DataGrid3.Columns(0)
Label10.Caption = Kon2.DataGrid3.Columns(3)
TextBox4.Text = Kon2.DataGrid3.Columns(3)
Label11.Caption = Kon2.DataGrid3.Columns(4)
TextBox5.Text = Kon2.DataGrid3.Columns(4)

If OptionButton1.Value = True Then
Label13.Caption = "Pagi"
End If
If OptionButton2.Value = True Then
Label13.Caption = "Malam"
End If
If OptionButton3.Value = True Then
OptionButton1.Value = False
Label13.Caption = "NS"
End If

End Sub

Private Sub UserForm_Initialize()

If App.PrevInstance = True Then
  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Aplikasi sudah dibuka", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)
      
End

End If

Module1.HideXCloseButton Me

Label6.Caption = Format(Now, "DD-MM-YYYY")
Label7.Caption = Format(Now, "hh")
Label9.Caption = Format(Now, "hh:mm:ss")

If (Label7 >= "7" Or Label7 <= "18") Then
OptionButton1 = True
OptionButton2 = False

ElseIf (Label7 >= "19" Or Label7 <= "24") Then
OptionButton1 = False
OptionButton2 = True

ElseIf (Label7 <= "6") Then
OptionButton1 = False
OptionButton2 = True

End If
'End If
'End If

With ComboBox1
.AddItem "Operator"
.AddItem "GL Produksi"
.AddItem "Teknisi"
.AddItem "Administrator"
End With

Label4.Caption = Format(Now, "dd")
Label5.Caption = Format(Now, "AMPM")

Jissk

'Masuk

End Sub

Sub Jissk()
If (Label4.Caption = "01" Or Label4.Caption = "16") And (Label5.Caption = "PM") Then
Shell App.Path & "\BackJissk.bat", vbMaximizedFocus
End If
End Sub

Sub Login()
If Label4.Caption = "01" Or "16" And Label5.Caption = "PM" Then
Shell App.Path & "\BackJissk.bat", vbMaximizedFocus
End If
End Sub

Sub Masuk()
Set kon = New ADODB.Connection
Set rs2 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\RecordLogin.mdb"
rs2.Open "select*from Log", kon, adOpenDynamic, adLockOptimistic

ListBox1.AddItem rs2!Shift
ListBox2.AddItem rs2!Tanggal
ListBox3.AddItem rs2!AMPM

rs2.Close

End Sub

Sub simpan()

Dim Waktu
If OptionButton1.Value = True Then
Waktu = "Pagi"
End If
If OptionButton2.Value = True Then
Waktu = "Malam"
End If

Set kon = New ADODB.Connection
Set rs2 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\RecordLogin.mdb"
Dim SQLTambah As String
SQLTambah = "insert into PIC (Tanggal,Jam,Pengguna,Shift,AMPM) values ('" & Label6 & "','" & Label9 & "','" & Label12 & "','" & Waktu & "','" & Label5 & "')"

kon.Execute SQLTambah

'rs2.Close
End Sub

Sub kriteria()

'NonShift
If (TextBox3.Text <> Label6.Caption And TextBox4.Text = Label13 And TextBox5.Text = Label5.Caption) Then
simpan

'Pagi -Malam
ElseIf (TextBox3.Text = Label6.Caption And TextBox4.Text <> Label13 And TextBox5.Text <> Label5.Caption) Then
simpan
'End If

'Malam -Pagi
ElseIf (TextBox3.Text <> Label6.Caption And TextBox4.Text <> Label13 And TextBox5.Text <> Label5.Caption) Then
simpan

'LompatHari
ElseIf (TextBox3.Text <> Label6.Caption And TextBox4.Text = Label13 And TextBox5.Text <> Label5.Caption) Then
simpan

End If
'End If
'End If
'End If
'End If

'NonShift
'If (Label6 <> Label10 And Label11 = Label5 And Label8 = Label3) Then
'simpan

'Pagi -Malam
'ElseIf (Label6 = Label10 And Label11 <> Label5 And Label8 <> Label3) Then
'simpan
'End If

'Malam -Pagi
'ElseIf (Label6 <> Label10 And Label11 <> Label5 And Label8 <> Label3) Then
'simpan



End Sub

