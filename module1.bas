Attribute VB_Name = "basMesh"
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Foll(50) As Byte
Public LeftN(50) As Byte
Public upperN(50) As Byte
Global Const Xmult = 0, Xcplx = 1, Xstrip = 2
Global Const Xlist = 3, Xfact = 4, Xphamp = 5
Global Const Xvar = 6, Xqfx = 7, Xcram = 8, Xrich = 9
Global Const PI = 3.141592654
Global ascs(10)
Global canvas() As Design
Type Coord
 X As Byte
 Y As Byte
 Terminal As Boolean
End Type
Public Const CELL_SPACING = 950

Public Function strCount(szSearch, szChar)
 searchpointer = InStr(szSearch, szChar)
 Ocurrances = 0
 Do Until searchpointer = 0
  Ocurrances = Ocurrances + 1
  searchpointer = InStr(searchpointer + 1, szSearch, szChar)
  DoEvents
 Loop
 strCount = Ocurrances
End Function
