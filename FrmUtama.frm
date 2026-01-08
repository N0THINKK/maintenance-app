VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Sheet Mesin AC90"
   ClientHeight    =   5595
   ClientLeft      =   30
   ClientTop       =   5385
   ClientWidth     =   6750
   ControlBox      =   0   'False
   Icon            =   "FrmUtama.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "FrmUtama.dsx":1084A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngFormWidth As Long
Private lngFormHeight As Long

Public kon As New ADODB.Connection
Public rs1 As New ADODB.Recordset

Public Bon As New ADODB.Connection
Public rs2 As New ADODB.Recordset

Private Sub Form_Load()
    Dim Ctl As Control
    lngFormWidth = ScaleWidth
    lngFormHeight = ScaleHeight
    On Error Resume Next
    For Each Ctl In Me
        Ctl.Tag = Ctl.Left & " " & Ctl.Top & " " & _
            Ctl.Width & " " & Ctl.Height & " "
            Ctl.Tag = Ctl.Tag & Ctl.FontSize & " "
    Next Ctl
    On Error GoTo 0
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
    ScaleX = ScaleWidth / lngFormWidth
    ScaleY = ScaleHeight / lngFormHeight
    On Error Resume Next
    For Each Ctl In Me
        TempVisible = Ctl.Visible
        Ctl.Visible = False
        StartPoz = 1
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
            Ctl.Move D(0) * ScaleX, D(1) * ScaleY, _
                D(2) * ScaleX, D(3) * ScaleY
            Ctl.Width = D(2) * ScaleX
            Ctl.Height = D(3) * ScaleY
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

Private Sub CheckBox1_Click()
If CheckBox1 = True Then
CrimpStandartFull.Show
Else
CrimpStandartFull.Hide
End If
End Sub

Private Sub CommandButton1_Click()

If NoMesin.Text16 = "A" Then
FrmPHOPA.Show
End If

If NoMesin.Text16 = "B" Then
FrmPHOPB.Show
End If

If NoMesin.Text16 = "DE" Then
FrmPHOPDE.Show
End If

End Sub

Private Sub CommandButton10_Click()
FrmRecord.Show
FrmUtama.Hide

End Sub

Private Sub CommandButton11_Click()
USER.Show
Unload Me
End Sub

Private Sub CommandButton12_Click()
FrmPHTEK.Show
End Sub

Private Sub CommandButton13_Click()
CommandButton13.Visible = False
End Sub

Private Sub CommandButton14_Click()
Jissk.Show
End Sub

Private Sub CommandButton15_Click()

PhAplA.Show

End Sub

Private Sub CommandButton16_Click()
Mikrometer.Show
End Sub

Private Sub CommandButton17_Click()
StripDepth.Show
End Sub

Private Sub CommandButton18_Click()
DailyReport.Show
End Sub

Private Sub CommandButton19_Click()
Abnormal.Show
End Sub

Private Sub CommandButton6_Click()
FrmHistry1.Show
End Sub

Private Sub CommandButton7_Click()
Unload Me

Frm_login.Show
End Sub

Private Sub CommandButton9_Click()

NoMesin.Show
NoMesin.Hide

If (FrmOP.Text34 = "") Or (FrmOP.Text34 Is Nothing) Then

  SetTimer hwnd, NV_CLOSEMSGBOX, 1000&, AddressOf TimerProc

  Call MessageBox(hwnd, "Barcode Kanban Terlebih dahulu", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)

Else

Shell App.Path & "\ConJissk2.bat", vbMinimizedFocus

Waktu
FrmOP.Show
     
End If

End Sub

Private Sub Label1_Click()
Label1.Caption = TextBox1.Text
End Sub

Private Sub TextBox1_Change()
TextBox1.Text = UCase(TextBox1.Text)
End Sub

Private Sub UserForm_Activate()

Module1.HideXCloseButton Me

Set kon = New ADODB.Connection
Set rs1 = New ADODB.Recordset
kon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\NmrMesin.mdb"
rs1.Open "select*from Nmr", kon, adOpenDynamic, adLockOptimistic

Login

'Label5.Caption = Kon2.DataGrid3.Columns(3)
'Label6.Caption = Kon2.DataGrid3.Columns(5)

Dim Berhenti As Boolean
Berhenti = False
Do Until Berhenti
Label2.Caption = Format(Now, "yyyy/mm/dd hh:mm:ss")
DoEvents
Loop

End Sub

Private Sub Form_Activate()

End Sub

Sub Waktu()
Dim time1, time2

time1 = Now
time2 = Now + TimeValue("0:00:03")
    Do Until time1 >= time2
        DoEvents
        time1 = Now()
    Loop
End Sub


Sub Login()
Set Bon = New ADODB.Connection
Set rs2 = New ADODB.Recordset
Bon.Open "provider=Microsoft.jet.oledb.4.0;data source=" & App.Path & "\RecordLogin.mdb"
rs2.Open "select*from PIC", Bon, adOpenDynamic, adLockOptimistic

'ListBox1.AddItem rs2!Shift
'ListBox2.AddItem rs2!NoUrut

rs2.Close

End Sub

Private Sub UserForm_Initialize()
If App.PrevInstance = True Then
  SetTimer hwnd, NV_CLOSEMSGBOX, 800&, AddressOf TimerProc

  Call MessageBox(hwnd, "Aplikasi sudah dibuka", _
      MSG_TITLE, MB_ICONQUESTION Or MB_TASKMODAL)
      
      End
      
End If
End Sub
