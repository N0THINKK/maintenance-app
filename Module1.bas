Attribute VB_Name = "Module1"
Option Explicit
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private hwnd As Long
Private lStyle As Long

#If VBA7 And Win64 Then
Private Declare PtrSafe Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongLong
Private Declare PtrSafe Function GetWindowLong Lib "user32" _
        Alias "GetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long) As LongLong
Private Declare PtrSafe Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As LongLong
Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As Long) As LongLong
Private Declare PtrSafe Function GetMenuItemCount Lib "user32" _
                (ByVal hMenu As Long) As LongLong
Private Declare PtrSafe Function GetSystemMenu Lib "user32" _
                (ByVal hwnd As Long, _
                ByVal bRevert As Long) As LongLong
Private Declare PtrSafe Function RemoveMenu Lib "user32" _
                (ByVal hMenu As Long, ByVal nPosition As Long, _
                ByVal wFlags As Long) As LongLong
#Else
Private Declare Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
                (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" _
                (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" _
                (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" _
                (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" _
                (ByVal hMenu As Long, ByVal nPosition As Long, _
                ByVal wFlags As Long) As Long
#End If

Sub HideXCloseButton(oForm As Object)
hwnd = FindWindow("ThunderDFrame", oForm.Caption)
lStyle = GetWindowLong(hwnd, GWL_STYLE)
SetWindowLong hwnd, GWL_STYLE, lStyle And Not WS_SYSMENU
End Sub

