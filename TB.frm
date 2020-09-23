VERSION 5.00
Begin VB.Form Fo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "TB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox SA 
      AutoRedraw      =   -1  'True
      Height          =   400
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   885
      TabIndex        =   1
      Top             =   1080
      Width           =   940
   End
   Begin VB.PictureBox SB 
      AutoRedraw      =   -1  'True
      Height          =   400
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   885
      TabIndex        =   0
      Top             =   600
      Width           =   940
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1800
   End
End
Attribute VB_Name = "Fo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
 Dim rct As RECT
 Dim Handle As Long, FindClass As Long
 ShowStartButton
 FindClass& = FindWindow("Shell_TrayWnd", "")
 Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
 GetWindowRect Handle&, rct
 DC& = GetDC(Handle&)
 BitBlt SB.hdc, 0, 0, 60, 25, DC&, 0, 0, SRCCOPY
 HideStartButton
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
 Static done
 Dim pnt As POINTAPI
 Dim Handle As Long, FindClass As Long
 Dim x As Long, y As Long
 FindClass& = FindWindow("Shell_TrayWnd", "")
 Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
 GetCursorPos pnt
 hW& = WindowFromPoint(pnt.x, pnt.y)
 DC& = GetDC(FindClass&)
 If hW = FindClass Or hW = Handle Then
  If done = 0 Then
    p1 = 1
    p2 = 10
    For t = 0 To 10
     For x = 0 To 58
      For y = 0 To 22
       pix1 = GetPixel(SB.hdc, x, y)
       pix2 = GetPixel(SA.hdc, x, y)
       UnRGB pix1, R1%, G1%, B1%
       UnRGB pix2, R2%, G2%, B2%
       Rr = (R1 * p1 + R2 * p2) / 11
       Gg = (G1 * p1 + G2 * p2) / 11
       Bb = (B1 * p1 + B2 * p2) / 11
       SetPixel DC&, x, y, RGB(Rr, Gg, Bb)
     Next y, x
     p1 = p1 + 1
     p2 = p2 - 1
    Next t
    done = 1
  End If
  ShowStartButton
 Else
  If done = 1 Then
    p1 = 10
    p2 = 1
    DC2& = GetDC(Handle&)
    BitBlt SB.hdc, 0, 0, 60, 25, DC2&, 0, 0, SRCCOPY
    For t = 0 To 10
     If t = 1 Then HideStartButton
     For x = 1 To 58
      For y = 1 To 22
       pix1 = GetPixel(SB.hdc, x, y)
       pix2 = GetPixel(SA.hdc, x, y)
       UnRGB pix1, R1%, G1%, B1%
       UnRGB pix2, R2%, G2%, B2%
       Rr = (R1 * p1 + R2 * p2) / 11
       Gg = (G1 * p1 + G2 * p2) / 11
       Bb = (B1 * p1 + B2 * p2) / 11
       SetPixel DC&, x, y, RGB(Rr, Gg, Bb)
     Next y, x
     p1 = p1 - 1
     p2 = p2 + 1
    Next t
    done = 0
  End If
 End If
 DeleteObject FindClass&
 DeleteObject Handle&
 DeleteObject DC&
 DeleteObject DC2&
End Sub
