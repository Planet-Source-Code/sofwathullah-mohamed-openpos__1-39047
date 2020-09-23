Attribute VB_Name = "Module1"
'// ===============================================================================//
'// Program: OpenPOS(Point of Sales)                                               //
'// Developed by: Sofwathullah Mohamed                                             //
'// Sofwath@Hotmail.Com                                                            //
'// You are free to use and modify this program as long as you give credit to the  //
'// original developer. Any comments or bugs report to sofwath@hotmail.com         //
'// Ver: 0.1                                                                       //
'// This Program is Still Under Development and Some of the Modules are Missing    //
'// ===============================================================================//

Option Explicit

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()

Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" _
         (ByVal wReturnCode As Long, ByVal dwReserved As Long) As Long

Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Declare Function WaitForSingleObject Lib "kernel32" _
         (ByVal hHandle As Long, _
         ByVal dwMilliseconds As Long) As Long
         
Declare Function PostMessage Lib "user32" _
    Alias "PostMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Declare Function IsWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hnd As Long, ByVal clval As Long, ByVal alph As Byte, ByVal flago As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias _
  "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&

'Constants used by the API functions
Public Const WM_CLOSE = &H10
Public Const INFINITE = &HFFFFFFFF
Public indexSales As Integer


Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Public Sub FormMove(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hwnd, &HA1, 2, 0&)
End Sub
Public Sub MakeChildTrans(transP As Integer, frm As Object)
    Dim lOldStyle As Long
    Dim bTrans As Byte ' The level of transparency (0 - 255)

    bTrans = transP
    lOldStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
    SetWindowLong frm.hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hwnd, 0, bTrans, LWA_ALPHA
End Sub
Public Sub DoMsg(msg As String)
    frmMessage.Label1.Caption = msg
    frmMessage.Show 1
    frmMessage.Timer1.Enabled = True
End Sub
