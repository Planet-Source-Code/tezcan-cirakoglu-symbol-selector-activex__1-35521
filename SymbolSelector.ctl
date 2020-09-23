VERSION 5.00
Begin VB.UserControl SymbolSelector 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   170
   ToolboxBitmap   =   "SymbolSelector.ctx":0000
End
Attribute VB_Name = "SymbolSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************
' SymbolSelector.ocx (It self-explains what it does :)
' Feel free to use it in educational, commercial or whatever apps
' No need to give any link to me in your app, or code
' It' totally free...
' tezcan_cirakoglu@hotmail.com
' ****************************************************************

' One more thing, you should catch that this control is
' a kind of implementation of another submission on psc.
' It was Color Selector, so if you want, you can thank to
' this author

' By the way, i'm so sorry that i don't use comments in my code

'Private Variables
Private IsButDown As Boolean
Private IsInFocus As Boolean
Private RBut As RECT

'Public Enums
Public Enum ssAppearanceConstants
    Flat
    [3D]
End Enum

'Private Constants
Private Const m_def_Appearance = ssAppearanceConstants.[3D]
Private Const m_def_BackColor = &H80000005
Private Const m_def_SymbolFont = "Symbol"
Private Const m_def_SymbolID = 0

'Private Property Variables
Private m_BackColor           As OLE_COLOR
Private m_Appearance          As ssAppearanceConstants
Private m_SymbolFont          As String
Private m_SymbolID            As Integer

Public Event SelectionChange(FontName As String, Item As Integer)

'*******************************************************************
'USER CONTROL
Private Sub UserControl_Initialize()
   ScaleMode = vbPixels
End Sub

Private Sub UserControl_GotFocus()
    IsInFocus = True
    Call RedrawControl
End Sub

Private Sub UserControl_LostFocus()
    IsInFocus = False
    Call RedrawControl
End Sub

Private Sub UserControl_InitProperties()
   m_Appearance = m_def_Appearance
   m_BackColor = m_def_BackColor
   m_SymbolFont = m_def_SymbolFont
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      If (X >= RBut.Left And X <= RBut.Right) And (Y >= RBut.Top And Y <= RBut.Bottom) Then
         IsButDown = True
         Call RedrawControl
      End If
   End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If IsButDown Then
      If Not ((X >= RBut.Left And X <= RBut.Right) And (Y >= RBut.Top And Y <= RBut.Bottom)) Then
         IsButDown = False
         Call RedrawControl
      End If
   End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      If IsButDown Then
         IsButDown = False
         Call RedrawControl
      End If
      
      If ((X >= ScaleLeft And X <= ScaleWidth) And (Y >= ScaleTop And Y <= ScaleHeight)) Then
         Call ShowSymbolPalette
      End If
   End If
End Sub

Private Sub UserControl_Paint()
   RedrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
   m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
   m_SymbolFont = PropBag.ReadProperty("SymbolFont", m_def_SymbolFont)
   m_SymbolID = PropBag.ReadProperty("SymbolID", m_def_SymbolID)
   RedrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
   Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
   Call PropBag.WriteProperty("SymbolFont", m_SymbolFont, m_def_SymbolFont)
   Call PropBag.WriteProperty("SymbolID", m_SymbolID, m_def_SymbolID)
End Sub
'*******************************************************************

'*******************************************************************
'PROPERTIES
Public Property Get Appearance() As ssAppearanceConstants
   Appearance = m_Appearance
End Property

Public Property Let Appearance(NewVal As ssAppearanceConstants)
   m_Appearance = NewVal
   PropertyChanged "Appearance"
   Call RedrawControl
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   m_BackColor = New_BackColor
   PropertyChanged "BackColor"
   Call RedrawControl
End Property

Public Property Get SymbolFont() As String
   SymbolFont = m_SymbolFont
End Property

Public Property Let SymbolFont(NewVal As String)
   m_SymbolFont = NewVal
   UserControl.Font = m_SymbolFont
   PropertyChanged "SymbolFont"
   Call RedrawControl
End Property

Public Property Get SymbolID() As Integer
   SymbolID = m_SymbolID
End Property

Public Property Let SymbolID(NewID As Integer)
   m_SymbolID = NewID
   PropertyChanged ("SymbolID")
   Call RedrawControl
End Property

Public Property Get SelectedItem() As Integer
   SelectedItem = m_SymbolID
End Property

'*******************************************************************

'*******************************************************************
'Private Functions
Private Sub RedrawControl()
    Dim Rct As RECT
    Dim Brsh As Long, Clr As Long
    Dim CurFont As String
    Dim lx As Long, ty As Long
    Dim rx As Long, by As Long
    
    lx = ScaleLeft: ty = ScaleTop
    rx = ScaleWidth: by = ScaleHeight
    
    Cls
    
    'Draw background
    Call SetRect(Rct, 0, 0, rx, by)
    Call OleTranslateColor(m_BackColor, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FillRect(hdc, Rct, Brsh)
    If m_Appearance = [3D] Then
        Call DrawEdge(hdc, Rct, EDGE_SUNKEN, BF_RECT)
    Else
        Call DrawEdge(hdc, Rct, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT Or BF_MONO)
    End If
    Call DeleteObject(Brsh)
    Call DeleteObject(Clr)
    
    'Draw button
    CurFont = UserControl.FontName
    UserControl.FontName = "Marlett"
    Call OleTranslateColor(vbButtonFace, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    If m_Appearance = [3D] Then
        If IsButDown Then
            Call SetRect(RBut, rx - 15, 2, rx - 2, by - 2)
            Call FillRect(hdc, RBut, Brsh)
            Call DrawEdge(hdc, RBut, EDGE_RAISED, BF_RECT Or BF_FLAT)
            Call SetRect(Rct, RBut.Left + 2, RBut.Top, RBut.Right, RBut.Bottom)
            Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        Else
            Call SetRect(RBut, rx - 15, 2, rx - 2, by - 2)
            Call FillRect(hdc, RBut, Brsh)
            Call DrawEdge(hdc, RBut, EDGE_RAISED, BF_RECT)
            Call SetRect(Rct, RBut.Left, RBut.Top, RBut.Right, RBut.Bottom - 1)
            Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        End If
    Else
        Call SetRect(RBut, rx - 15, ty, rx, by)
        Call FillRect(hdc, RBut, Brsh)
        Call DrawEdge(hdc, RBut, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT)
        Call SetRect(Rct, RBut.Left + 1, RBut.Top, RBut.Right, RBut.Bottom - 1)
        Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    End If
    UserControl.FontName = m_SymbolFont
    If m_SymbolID > -1 Then
      Call SetRect(Rct, 0, 0, ScaleWidth - 15, ScaleHeight)
      Call DrawText(hdc, Chr(m_SymbolID), 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      RaiseEvent SelectionChange(m_SymbolFont, m_SymbolID)
    End If
    If IsInFocus Then
      Rct.Top = Rct.Top + 3
      Rct.Left = Rct.Left + 3
      Rct.Bottom = Rct.Bottom - 3
      Rct.Right = Rct.Right - 3
      Call DrawFocusRect(hdc, Rct)
    End If
    Call DeleteObject(Brsh)
    Call DeleteObject(Clr)
End Sub

Private Sub ShowSymbolPalette()
   Dim ClrCtrlPos As RECT, Rc As RECT
   
   Call GetWindowRect(hwnd, ClrCtrlPos)
   m_FontName = m_SymbolFont
   Load frmSymbol
   With frmSymbol
      
      .Left = ClrCtrlPos.Left * Screen.TwipsPerPixelX
      .Top = ClrCtrlPos.Bottom * Screen.TwipsPerPixelY
      
      If (.Top + .Height) > Screen.Height Then
         .Top = ClrCtrlPos.Top * Screen.TwipsPerPixelY - .Height
      End If
      
      .Show vbModal
      
      SymbolID = .SelectedItem
      
      If Not .IsCanceled Then
         Call RedrawControl
      End If
   
   End With
   Unload frmSymbol
End Sub
'*******************************************************************
