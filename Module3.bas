Attribute VB_Name = "Module3"
Option Explicit
Public Const WM_DISPLAYCHANGE = &H7E
Public Const HWND_BROADCAST = &HFFFF&
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const BITSPIXEL = 12
Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public OldX As Long, OldY As Long, nDC As Long
Public Function ChangeRes(X As Long, Y As Long, Bits As Long)
    Dim DevM As DEVMODE, ScInfo As Long, erg As Long, an As VbMsgBoxResult
    'Tambahkan info ke DevM
    erg = EnumDisplaySettings(0&, 0&, DevM)
    'ini yang berfungsi untuk mengubah
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = X 'lebar
    DevM.dmPelsHeight = Y 'tinggi
    DevM.dmBitsPerPel = Bits '(dapat 8, 16, 24, 32 atau bahkan 4)
    'sekarang, mengubah layar jika memungkinkan
    erg = ChangeDisplaySettings(DevM, CDS_TEST)
    'cek jika berhasil
    Select Case erg&
        Case DISP_CHANGE_RESTART
            an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
            If an = vbYes Then
                erg& = ExitWindowsEx(EWX_REBOOT, 0&)
            End If
        Case DISP_CHANGE_SUCCESSFUL
            erg = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
            ScInfo = Y * 2 ^ 16 + X
    End Select
End Function
