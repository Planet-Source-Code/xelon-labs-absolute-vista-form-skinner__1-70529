VERSION 5.00
Begin VB.UserControl Glasskinner 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape Shape 
      BackColor       =   &H0059341C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   8
      FillColor       =   &H00F2E2D9&
      FillStyle       =   7  'Diagonal Cross
      Height          =   2415
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Glasskinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private myform As Form

Private Sub Timer1_Timer()

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
assign
End Sub

Sub assign()
On Error Resume Next
    Set myform = UserControl.Parent
Set Skin.mfrm = myform
myform.BorderStyle = 0
Skin.Show
settings True, myform.MaxButton, myform.MinButton, True, True, myform.Moveable
End Sub

Public Sub settings(closeb As Boolean, maxb As Boolean, minb As Boolean, resb As Boolean, transb As Boolean, movb As Boolean)
Skin.Topp.Visible = movb
Skin.Max.Visible = maxb
Skin.Min.Visible = minb
Skin.Topp.Visible = transb
Skin.Rightp.Visible = resb
Skin.Leftp.Visible = resb
Skin.Bottomp.Visible = resb
Skin.BRp.Visible = resb
Skin.blp.Visible = resb
Skin.tlp.Visible = resb
Skin.cloz.Visible = closeb
Skin.trans = transb
End Sub

Private Sub UserControl_Resize()
Shape.Width = Width
Shape.Height = Height
End Sub
