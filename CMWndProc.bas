Attribute VB_Name = "CMWndProc"
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long

Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_COMMAND = &H111
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_GETFONT = &H31
Private Const WM_MEASUREITEM = &H2C
Private Const WM_NCHITTEST = &H84
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_WININICHANGE = &H1A
Private Const WM_SETCURSOR = &H20
Private Const WM_SETTINGCHANGE = WM_WININICHANGE

Private m_CoolMenuObj As CoolMenu

Public Property Set CoolMenuObj(ByVal vData As CoolMenu)
    Set m_CoolMenuObj = vData
End Property

Public Property Get CoolMenuObj() As CoolMenu
    Set CoolMenuObj = m_CoolMenuObj
End Property


Public Function WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
  On Error GoTo ErrorHandle
  
  Select Case msg&
  
    'All other info are dynamic (I hope)
    Case WM_SETTINGCHANGE: Call m_CoolMenuObj.GetMenuFont(True)
    Case WM_SYSCOLORCHANGE: Call m_CoolMenuObj.GetMenuFont(True)
    
    Case WM_MEASUREITEM:

            If m_CoolMenuObj.OnMeasureItem(lparam&) Then
              WindowProc = True
              Exit Function
            End If

            
    Case WM_DRAWITEM:

            If m_CoolMenuObj.OnDrawItem(lparam&) Then
              WindowProc = True
              Exit Function
            End If

    
    Case WM_INITMENUPOPUP:

            Call CallWindowProc(m_CoolMenuObj.PrevWndProc, ByVal hwnd&, ByVal msg&, ByVal wparam&, ByVal lparam&)
            Call m_CoolMenuObj.OnInitMenuPopup(wparam&, LoWord(lparam&), CBool(HiWord(lparam&)))
            WindowProc = 0&
            Exit Function
            
    Case WM_MENUCHAR:
            
            Dim result As Long
            result = m_CoolMenuObj.OnMenuChar(LoWord(wparam&), HiWord(wparam&), lparam&)

            If result <> 0 Then
              WindowProc = result
              Exit Function
            End If
            
    Case WM_MENUSELECT:
            
            Call m_CoolMenuObj.OnMenuSelect(LoWord(wparam&), HiWord(wparam&), lparam&)
      
  End Select
  
Continue:
  WindowProc& = CallWindowProc(m_CoolMenuObj.PrevWndProc, hwnd&, msg&, wparam&, lparam&)
  Exit Function
  
ErrorHandle:
  Debug.Print Err.Number; Err.Description
  Err.Clear
'  GoTo Continue
  m_CoolMenuObj.Install 0&
End Function


Public Function HiWord(LongIn As Long) As Integer
     HiWord% = (LongIn& And &HFFFF0000) \ &H10000
End Function

Public Function LoWord(LongIn As Long) As Integer
  Dim l As Long
  
  l& = LongIn& And &HFFFF&
  
  If l& > &H7FFF Then
       LoWord% = l& - &H10000
  Else
       LoWord% = l&
  End If
End Function

Public Function HiByte(WordIn As Integer) As Byte
  
  If WordIn% And &H8000 Then
    HiByte = &H80 Or ((WordIn% And &H7FFF) \ &HFF)
  Else
    HiByte = WordIn% \ 256
  End If

End Function

Public Function LoByte(WordIn As Integer) As Byte
  LoByte = WordIn% And &HFF&
End Function

Public Function MakeLong(LoWord As Integer, HiWord As Integer) As Long
'Useful when converting code from C++

  Dim nLoWord As Long
  
  If LoWord% < 0 Then
    nLoWord& = LoWord% + &H10000
  Else
    nLoWord& = LoWord%
  End If

  MakeLong& = CLng(nLoWord&) Or (HiWord% * &H10000)
End Function

Public Function MakeWord(LoByte As Byte, HiByte As Byte) As Integer
'Useful when converting code from C++
  Dim nLoByte As Integer

  If LoByte < 0 Then
    nLoByte = LoByte + &H100
  Else
    nLoByte = LoByte
  End If

  MakeWord = CInt(nLoByte) Or (HiByte * &H100)
End Function

