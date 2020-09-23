VERSION 5.00
Begin VB.Form frmSymbol 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsCanceled As Boolean
Public SelectedItem As Integer

Private R(1 To 255) As RECT

Private LastButId As Integer
Private MouseButDown As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyEscape) Then
      Me.Hide
   End If
End Sub

Private Sub Form_Load()
   Dim CurLeft As Integer, CurTop As Integer, Rc As RECT
   Me.Width = ((18 * 15) * Screen.TwipsPerPixelX) - 140
   Me.Height = ((18 * 17) * Screen.TwipsPerPixelY) - 140
   Me.FontName = m_FontName
   CurTop = 2: CurLeft = 2
   IsCanceled = False
   For i = 1 To 255
      If (i Mod 15) > 0 Then
         Call SetRect(R(i), CurLeft, CurTop, CurLeft + 16, CurTop + 16)
         CurLeft = CurLeft + 18
      Else
         CurLeft = 2
         CurTop = CurTop + 16
         Call SetRect(R(i), CurLeft, CurTop, CurLeft + 16, CurTop + 16)
      End If
   Next
   
   Call SetRect(Rc, 0, 0, ScaleWidth, ScaleHeight)
   Call DrawEdge(hdc, Rc, BDR_RAISEDINNER, BF_RECT)
   
   Call SetCapture(hwnd)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      MouseButDown = True
      For i = 1 To 255
         If PtInRect(R(i), X, Y) Then
            Call DrawEdge(hdc, R(i), BDR_SUNKENOUTER, BF_RECT)
         Else
            Call DrawEdge(hdc, R(i), BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
         End If
      Next
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Integer, IsMouseOnBut As Boolean
   For i = 1 To 255
      If PtInRect(R(i), X, Y) Then
         If MouseButDown Then
            Call DrawEdge(hdc, R(i), BDR_SUNKENOUTER, BF_RECT)
         Else
            Call DrawEdge(hdc, R(i), BDR_RAISEDINNER, BF_RECT)
         End If
      Else
         Call DrawEdge(hdc, R(i), BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
      End If
   Next
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim IsMouseOver As Boolean
   IsMouseOver = X >= 0 And Y >= 0 And X <= ScaleWidth And Y <= ScaleHeight
   If IsMouseOver Then
      For i = 1 To 255
         If PtInRect(R(i), X, Y) Then
            SelectedItem = i
            Call ReleaseCapture
            Call Form_KeyDown(vbKeyEscape, 0)
            IsCanceled = False
            Exit For
         End If
      Next
   Else
      Call ReleaseCapture
      IsCanceled = True
      Call Form_KeyDown(vbKeyEscape, 0)
   End If
   MouseButDown = False
End Sub

Private Sub Form_Paint()
   Dim i As Integer
   For i = 1 To 255
      Call DrawText(Me.hdc, Chr(i), 1&, R(i), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
   Next
End Sub
