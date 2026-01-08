Attribute VB_Name = "Module2"
  Option Explicit
  ' demo project showing how to use the API to manipulate a messagebox
  ' by Bryan Stafford of New Vision Software® - newvision@imt.net
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.
  
 Public Const MSG_TITLE = "Informasi"
  
  ' the max length of a path for the system (usually 260 or there abouts)
  ' this is used to size the buffer string for retrieving the class name of the active window below
  Public Const MAX_PATH As Long = 260&

  Public Const API_TRUE As Long = 1&
  Public Const API_FALSE As Long = 0&
  
  ' font *borrowed* from the form used to replace MessageBox font
  Public g_hBoldFont As Long
  
  Public Const MSGBOXTEXT As String = "Have you ever seen a standard message box with a different font than all the others on the system?"
  Public Const WM_SETFONT As Long = &H30

  ' made up constants for setting our timer
  Public Const NV_CLOSEMSGBOX As Long = &H5000&
  Public Const NV_MOVEMSGBOX As Long = &H5001&
  Public Const NV_MSGBOXCHNGFONT As Long = &H5002&

  ' MessageBox() Flags
  Public Const MB_ICONQUESTION As Long = &H20&
  Public Const MB_TASKMODAL As Long = &H2000&

  ' SetWindowPos Flags
  Public Const SWP_NOSIZE As Long = &H1&
  Public Const SWP_NOZORDER As Long = &H4&
  Public Const HWND_TOP As Long = 0&

  Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  ' API declares
  Public Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock&)
  
  Public Declare Function GetActiveWindow& Lib "user32" ()
  
  Public Declare Function GetDesktopWindow& Lib "user32" ()
  
  Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, _
                                                                        ByVal lpWindowName$)

  Public Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, _
                              ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$)

  Public Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal _
                                                        wMsg&, ByVal wParam&, lParam As Any)

  Public Declare Function MoveWindow& Lib "user32" (ByVal hwnd&, ByVal x&, ByVal y&, _
                                              ByVal nWidth&, ByVal nHeight&, ByVal bRepaint&)

  Public Declare Function ScreenToClientLong& Lib "user32" Alias "ScreenToClient" (ByVal hwnd&, _
                                                                                    lpPoint&)
  
  Public Declare Function GetDC& Lib "user32" (ByVal hwnd&)
  Public Declare Function ReleaseDC& Lib "user32" (ByVal hwnd&, ByVal hDC&)

  ' drawtext flags
  Public Const DT_WORDBREAK As Long = &H10&
  Public Const DT_CALCRECT As Long = &H400&
  Public Const DT_EDITCONTROL As Long = &H2000&
  Public Const DT_END_ELLIPSIS As Long = &H8000&
  Public Const DT_MODIFYSTRING As Long = &H10000
  Public Const DT_PATH_ELLIPSIS As Long = &H4000&
  Public Const DT_RTLREADING As Long = &H20000
  Public Const DT_WORD_ELLIPSIS As Long = &H40000
  
  Public Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hDC&, ByVal lpsz$, _
                                          ByVal cchText&, lpRect As RECT, ByVal dwDTFormat&)
  
  Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd&)
  
  Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, _
                                                        ByVal lpClassName$, ByVal nMaxCount&)

  Public Declare Function GetWindowRect& Lib "user32" (ByVal hwnd&, lpRect As RECT)
  
  Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, _
                                      ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
                                      
  Public Declare Function MessageBox& Lib "user32" Alias "MessageBoxA" (ByVal hwnd&, _
                                                ByVal lpText$, ByVal lpCaption$, ByVal wType&)

  Public Declare Function SetTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, _
                                                                            ByVal lpTimerFunc&)
  
  Public Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)

Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
  ' this is a callback function.  This means that windows "calls back" to this function
  ' when it's time for the timer event to fire
  
  ' first thing we do is kill the timer so that no other timer events will fire
  KillTimer hwnd, idEvent
  
  ' select the type of manipulation that we want to perform
  Select Case idEvent
    Case NV_CLOSEMSGBOX '<-- we want to close this messagebox after 4 seconds
      Dim hMessageBox&
      
      ' find the messagebox window
      hMessageBox = FindWindow("#32770", MSG_TITLE)
      
      ' if we found it make sure it has the keyboard focus and then send it an enter to dismiss it
      If hMessageBox Then
        Call SetForegroundWindow(hMessageBox)
        SendKeysA vbKeyReturn, True
      End If
      
    Case NV_MOVEMSGBOX '<-- we want to move this messagebox
      Dim hMsgBox&, xPoint&, yPoint&
      Dim stMsgBoxRect As RECT, stParentRect As RECT
      

      hMsgBox = FindWindow("#32770", "Position A Message Box")
    
      If hMsgBox Then
        Call GetWindowRect(hMsgBox, stMsgBoxRect)
        Call GetWindowRect(hwnd, stParentRect)
        
        xPoint = stParentRect.Left + (((stParentRect.Right - stParentRect.Left) \ 2) - _
                                              ((stMsgBoxRect.Right - stMsgBoxRect.Left) \ 2))
        yPoint = stParentRect.Top + (((stParentRect.Bottom - stParentRect.Top) \ 2) - _
                                              ((stMsgBoxRect.Bottom - stMsgBoxRect.Top) \ 2))
        
        If xPoint < 0 Then xPoint = 0
        If yPoint < 0 Then yPoint = 0
        If (xPoint + (stMsgBoxRect.Right - stMsgBoxRect.Left)) > (Screen.Width \ Screen.TwipsPerPixelX) Then
          xPoint = (Screen.Width \ Screen.TwipsPerPixelX) - (stMsgBoxRect.Right - stMsgBoxRect.Left)
        End If
        If (yPoint + (stMsgBoxRect.Bottom - stMsgBoxRect.Top)) > (Screen.Height \ Screen.TwipsPerPixelY) Then
          yPoint = (Screen.Height \ Screen.TwipsPerPixelY) - (stMsgBoxRect.Bottom - stMsgBoxRect.Top)
        End If
        
        
        Call SetWindowPos(hMsgBox, HWND_TOP, xPoint, yPoint, API_FALSE, API_FALSE, SWP_NOZORDER Or SWP_NOSIZE)
      End If
      
      Call LockWindowUpdate(API_FALSE)
      
      
    Case NV_MSGBOXCHNGFONT
      
      hMsgBox = FindWindow("#32770", "Change The Message Box Font")
    
      If hMsgBox Then
        Dim hStatic&, hButton&, stMsgBoxRect2 As RECT
        Dim stStaticRect As RECT, stButtonRect As RECT
        
        hStatic = FindWindowEx(hMsgBox, API_FALSE, "Static", MSGBOXTEXT)
        hButton = FindWindowEx(hMsgBox, API_FALSE, "Button", "OK")
        
        If hStatic Then
          Call GetWindowRect(hMsgBox, stMsgBoxRect2)
          Call GetWindowRect(hStatic, stStaticRect)
          Call GetWindowRect(hButton, stButtonRect)
          
          Call SendMessage(hStatic, WM_SETFONT, g_hBoldFont, ByVal API_TRUE)
          
          With stStaticRect
            Call ScreenToClientLong(hMsgBox, .Left)
            Call ScreenToClientLong(hMsgBox, .Right)
            
            Dim nRectHeight&, nHeightDifference&, hStaticDC&
            nHeightDifference = .Bottom - .Top
            
            hStaticDC = GetDC(hStatic)
            nRectHeight = DrawText(hStaticDC, MSGBOXTEXT, (-1&), stStaticRect, DT_CALCRECT Or DT_EDITCONTROL Or DT_WORDBREAK)
            
            Call ReleaseDC(hStatic, hStaticDC)
            nHeightDifference = nRectHeight - nHeightDifference
            
            Call MoveWindow(hStatic, .Left, .Top, .Right - .Left, nRectHeight, API_TRUE)
          End With
            
          With stButtonRect
            Call ScreenToClientLong(hMsgBox, .Left)
            Call ScreenToClientLong(hMsgBox, .Right)
            Call MoveWindow(hButton, .Left, .Top + nHeightDifference, .Right - .Left, .Bottom - .Top, API_TRUE)
          End With
          
          With stMsgBoxRect2
            Call MoveWindow(hMsgBox, .Left, .Top - (nHeightDifference \ 2), .Right - .Left, (.Bottom - .Top) + nHeightDifference, API_TRUE)
          
          End With
        End If
      End If
      Call LockWindowUpdate(API_FALSE)
  
  End Select
  
End Sub




