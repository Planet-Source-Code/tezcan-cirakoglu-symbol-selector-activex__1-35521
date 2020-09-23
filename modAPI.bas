Attribute VB_Name = "modAPI"
Public m_FontName As String

Public Const OPAQUE = 2
Public Const TRANSPARENT = 1
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_BOTTOM = &H8
Public Const BF_FLAT = &H4000      ' For flat rather than 3D borders
Public Const BF_LEFT = &H1
Public Const BF_MONO = &H8000      ' For monochrome borders.
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Public Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const DT_CALCRECT = &H400

Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hbitmap As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Option Explicit
Option Base 1


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public CustClrs() As OLE_COLOR
Public LastSavedCustClr As Integer

Public DefClr As OLE_COLOR
Public CurClr As OLE_COLOR

Public DefCap As String
Public MorCap As String

Public ShwDef As Boolean
Public ShwCus As Boolean
Public ShwMor As Boolean
Public ShwSys As Boolean
Public ShwTip As Boolean

Public Sub Timer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Call frmColorPalette.TipTimer(hwnd, uMsg, idEvent, dwTime)
End Sub


